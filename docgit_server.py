import os
import sys
import subprocess
import re
import shlex
import json
import uuid
import urllib.parse
import difflib
from flask import Flask, request, jsonify, send_from_directory, render_template_string
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ── Paths ────────────────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
    MEIPASS_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    MEIPASS_DIR = BASE_DIR

STATIC_DIR   = os.path.join(MEIPASS_DIR, 'static')
DOCGIT_EXE   = os.path.join(BASE_DIR, 'docgit.exe')
if not os.path.exists(DOCGIT_EXE):
    DOCGIT_EXE = None   # dev mode — use docgit.py

ANSI_ESCAPE  = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

# In-memory diff session store  {session_id: html_string}
DIFF_SESSIONS: dict = {}


# ── Helpers ──────────────────────────────────────────────────────────────────
def get_cmd_base():
    if DOCGIT_EXE:
        return [DOCGIT_EXE]
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'docgit.py')
    return [sys.executable, script]


def normalize_path(filepath):
    if not filepath:
        return filepath
    if filepath.startswith("file:///"):
        filepath = urllib.parse.unquote(filepath[8:])
    return filepath.replace("/", "\\")


def load_index(doc_dir):
    index_path = os.path.join(doc_dir, '.docgit', 'index.json')
    if not os.path.exists(index_path):
        return None
    with open(index_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def extract_paragraphs(docx_path):
    """Extract all content from a .docx as a list of strings (paragraphs, tables, images)."""
    try:
        from docx import Document
        from docx.oxml.ns import qn
        doc = Document(docx_path)
        items = []
        for element in doc.element.body:
            raw_tag = element.tag.split('}')[1] if '}' in element.tag else element.tag

            if raw_tag == 'p':  # paragraph
                text = ''.join(
                    node.text for node in element.iter(qn('w:t')) if node.text
                )
                drawings = element.findall('.//' + qn('w:drawing'))
                picts    = element.findall('.//' + qn('w:pict'))
                img_count = len(drawings) + len(picts)
                if img_count:
                    items.append(f"[Image{'s' if img_count > 1 else ''} x{img_count}]{' ' + text if text else ''}")
                else:
                    items.append(text)  # may be empty string — preserves blank lines

            elif raw_tag == 'tbl':  # table
                items.append('[TABLE]')
                for tr in element.findall('.//' + qn('w:tr')):
                    cells = []
                    for tc in tr.findall(qn('w:tc')):
                        cell_text = ''.join(
                            node.text for node in tc.iter(qn('w:t')) if node.text
                        )
                        cells.append(cell_text)
                    items.append('| ' + ' | '.join(cells) + ' |')
                items.append('[/TABLE]')

        return items
    except Exception as e:
        return [f'[Error reading file: {e}]']



def build_diff_html(left_paras, right_paras, left_label, right_label):
    """Build a self-contained side-by-side HTML diff page."""
    differ = difflib.SequenceMatcher(None, left_paras, right_paras)
    opcodes = differ.get_opcodes()

    left_rows  = []
    right_rows = []

    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            for k in range(i2 - i1):
                text = left_paras[i1 + k] or '&nbsp;'
                left_rows.append(f'<div class="line eq">{_esc(text)}</div>')
                right_rows.append(f'<div class="line eq">{_esc(text)}</div>')
        elif tag == 'replace':
            left_lines  = left_paras[i1:i2]
            right_lines = right_paras[j1:j2]
            max_len = max(len(left_lines), len(right_lines))
            for k in range(max_len):
                l = left_lines[k]  if k < len(left_lines)  else ''
                r = right_lines[k] if k < len(right_lines) else ''
                lh, rh = _inline_diff(l, r)
                left_rows.append(f'<div class="line del">{lh}</div>')
                right_rows.append(f'<div class="line ins">{rh}</div>')
        elif tag == 'delete':
            for i in range(i1, i2):
                left_rows.append(f'<div class="line del">{_esc(left_paras[i])}</div>')
                right_rows.append('<div class="line empty">&nbsp;</div>')
        elif tag == 'insert':
            for j in range(j1, j2):
                left_rows.append('<div class="line empty">&nbsp;</div>')
                right_rows.append(f'<div class="line ins">{_esc(right_paras[j])}</div>')

    left_html  = '\n'.join(left_rows)  or '<div class="line eq">(empty)</div>'
    right_html = '\n'.join(right_rows) or '<div class="line eq">(empty)</div>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>DocGit Diff</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', sans-serif; background: #1e1e2e; color: #cdd6f4; }}
  header {{ background: #181825; padding: 16px 24px; display: flex; align-items: center; gap: 12px; border-bottom: 1px solid #313244; }}
  header h1 {{ font-size: 20px; color: #89b4fa; }}
  header .legend {{ display: flex; gap: 16px; font-size: 12px; margin-left: auto; }}
  header .legend span {{ display: flex; align-items: center; gap: 6px; }}
  .swatch {{ width: 12px; height: 12px; border-radius: 2px; }}
  .sw-del {{ background: #f38ba8; }} .sw-ins {{ background: #a6e3a1; }} .sw-eq {{ background: #6c7086; }}
  .diff-container {{ display: flex; height: calc(100vh - 57px); }}
  .pane {{ flex: 1; overflow-y: auto; padding: 0; border-right: 1px solid #313244; }}
  .pane:last-child {{ border-right: none; }}
  .pane-header {{ position: sticky; top: 0; background: #181825; padding: 8px 16px; font-size: 13px; font-weight: 600; color: #89b4fa; border-bottom: 1px solid #313244; }}
  .line {{ padding: 3px 16px; font-size: 13px; font-family: 'Consolas', monospace; white-space: pre-wrap; word-break: break-word; min-height: 22px; }}
  .line.eq {{ color: #a6adc8; }}
  .line.del {{ background: rgba(243,139,168,0.15); color: #f38ba8; }}
  .line.ins {{ background: rgba(166,227,161,0.15); color: #a6e3a1; }}
  .line.empty {{ background: rgba(49,50,68,0.3); }}
  mark.del {{ background: rgba(243,139,168,0.4); color: #f38ba8; border-radius: 2px; }}
  mark.ins {{ background: rgba(166,227,161,0.4); color: #a6e3a1; border-radius: 2px; }}
</style>
</head>
<body>
<header>
  <h1>DocGit Diff</h1>
  <div class="legend">
    <span><span class="swatch sw-del"></span> Removed</span>
    <span><span class="swatch sw-ins"></span> Added</span>
    <span><span class="swatch sw-eq"></span> Unchanged</span>
  </div>
</header>
<div class="diff-container">
  <div class="pane">
    <div class="pane-header">{_esc(left_label)}</div>
    {left_html}
  </div>
  <div class="pane">
    <div class="pane-header">{_esc(right_label)}</div>
    {right_html}
  </div>
</div>
</body>
</html>"""


def _esc(text):
    """HTML-escape text."""
    return (str(text)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


def _inline_diff(a, b):
    """Return HTML for character-level inline diff of two strings."""
    sm = difflib.SequenceMatcher(None, list(a), list(b))
    left, right = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            chunk = _esc(''.join(a[i1:i2]))
            left.append(chunk); right.append(chunk)
        elif tag in ('replace', 'delete'):
            left.append(f'<mark class="del">{_esc("".join(a[i1:i2]))}</mark>')
        if tag in ('replace', 'insert'):
            right.append(f'<mark class="ins">{_esc("".join(b[j1:j2]))}</mark>')
    return ''.join(left), ''.join(right)


def run_subprocess(cmd, cwd, timeout=30):
    env = os.environ.copy()
    env['NO_COLOR'] = '1'
    env['TERM'] = 'dumb'
    try:
        r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                           env=env, timeout=timeout,
                           creationflags=subprocess.CREATE_NO_WINDOW)
        out = ANSI_ESCAPE.sub('', (r.stdout + r.stderr).strip())
        return out or f'(exited {r.returncode})', r.returncode
    except subprocess.TimeoutExpired:
        return 'Command timed out.', -1
    except Exception as e:
        return str(e), -1


# ── Static files ─────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory(STATIC_DIR, 'taskpane.html')


@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory(STATIC_DIR, filename)


# ── Diff viewer ───────────────────────────────────────────────────────────────
@app.route('/diff-view/<session_id>')
def diff_view(session_id):
    html = DIFF_SESSIONS.get(session_id)
    if not html:
        return "<h2>Diff session expired. Please run diff again.</h2>", 404
    return html


# ── API: branches ─────────────────────────────────────────────────────────────
@app.route('/api/branches', methods=['POST'])
def api_branches():
    data = request.get_json(force=True)
    filepath = normalize_path(data.get('filepath', ''))
    if not filepath:
        return jsonify({'error': 'No filepath'})
    doc_dir = os.path.dirname(filepath)
    index = load_index(doc_dir)
    if not index:
        return jsonify({'branches': [], 'current': 'main', 'initialized': False})
    branches = list(index.get('branches', {}).keys())
    current  = index.get('HEAD', 'main')
    return jsonify({'branches': branches, 'current': current, 'initialized': True})


# ── API: commits (for rollback dropdown) ─────────────────────────────────────
@app.route('/api/commits', methods=['POST'])
def api_commits():
    data = request.get_json(force=True)
    filepath = normalize_path(data.get('filepath', ''))
    if not filepath:
        return jsonify({'error': 'No filepath'})
    doc_dir = os.path.dirname(filepath)
    index   = load_index(doc_dir)
    if not index:
        return jsonify({'commits': []})
    branch    = index.get('HEAD', 'main')
    commit_id = index['branches'].get(branch)
    commits   = []
    while commit_id and commit_id in index['commits']:
        c = index['commits'][commit_id]
        commits.append({
            'id': commit_id,
            'message': c['message'],
            'timestamp': c['timestamp'][:16].replace('T', ' ')
        })
        commit_id = c.get('parent')
    return jsonify({'commits': commits, 'current_branch': branch})


# ── API: has-changes (for switch guard) ──────────────────────────────────────
@app.route('/api/has-changes', methods=['POST'])
def api_has_changes():
    """Returns {changed: bool, files: [list of modified/untracked files]}"""
    import hashlib
    data     = request.get_json(force=True)
    filepath = normalize_path(data.get('filepath', ''))
    if not filepath:
        return jsonify({'error': 'No filepath'})
    doc_dir = os.path.dirname(filepath)
    index   = load_index(doc_dir)
    if not index:
        return jsonify({'changed': False, 'files': []})

    branch    = index.get('HEAD', 'main')
    commit_id = index['branches'].get(branch)
    tree      = index['commits'][commit_id]['tree'] if commit_id and commit_id in index['commits'] else {}

    changed = []
    try:
        working = [f for f in os.listdir(doc_dir)
                   if f.endswith('.docx') and not f.startswith('~$') and os.path.isfile(os.path.join(doc_dir, f))]
        for f in working:
            full = os.path.join(doc_dir, f)
            hasher = hashlib.sha256()
            with open(full, 'rb') as fh:
                hasher.update(fh.read())
            h = hasher.hexdigest()
            if f not in tree:
                changed.append(f + ' (untracked)')
            elif tree[f] != h:
                changed.append(f + ' (modified)')
        for f in tree:
            if f not in working:
                changed.append(f + ' (deleted)')
    except Exception as e:
        return jsonify({'error': str(e)})

    return jsonify({'changed': len(changed) > 0, 'files': changed})



@app.route('/api/diff', methods=['POST'])
def api_diff():
    """
    Body:
      filepath   - active document
      mode       - "current"   : working file vs last commit
                   "branches"  : branchA tip vs branchB tip
                   "commits"   : commitA vs commitB
      branch_a, branch_b  (for mode=branches)
      commit_a, commit_b  (for mode=commits)
    Returns: { session_id, diff_url }
    """
    data     = request.get_json(force=True)
    filepath = normalize_path(data.get('filepath', ''))
    mode     = data.get('mode', 'current')
    if not filepath:
        return jsonify({'error': 'No filepath'})

    doc_dir = os.path.dirname(filepath)
    index   = load_index(doc_dir)
    if not index:
        return jsonify({'error': 'Not a docgit repository. Run init first.'})

    objects_dir = os.path.join(doc_dir, '.docgit', 'objects')
    doc_name    = os.path.basename(filepath)
    rel_path    = doc_name   # commit tree keys are relative paths

    def get_commit_docx(commit_id):
        """Return path to the docx stored in a commit, or None."""
        c = index['commits'].get(commit_id)
        if not c:
            return None
        tree = c.get('tree', {})
        # try exact match or match by filename
        h = tree.get(rel_path) or tree.get(filepath)
        if not h:
            # try any value that matches the filename
            for k, v in tree.items():
                if os.path.basename(k).lower() == doc_name.lower():
                    h = v; break
        if not h:
            return None
        p = os.path.join(objects_dir, f'{h}.docx')
        return p if os.path.exists(p) else None

    if mode == 'current':
        # Working file vs HEAD commit
        head_branch = index.get('HEAD', 'main')
        head_commit = index['branches'].get(head_branch)
        if not head_commit:
            return jsonify({'error': 'No commits yet. Nothing to diff against.'})
        committed_path = get_commit_docx(head_commit)
        if not committed_path:
            return jsonify({'error': f'File "{doc_name}" not found in last commit.'})
        left_paras  = extract_paragraphs(committed_path)
        right_paras = extract_paragraphs(filepath) if os.path.exists(filepath) else []
        left_label  = f'Last Commit ({head_commit[:7]}) — {doc_name}'
        right_label = f'Working Copy — {doc_name}'

    elif mode == 'branches':
        branch_a = data.get('branch_a', '')
        branch_b = data.get('branch_b', '')
        if not branch_a or not branch_b:
            return jsonify({'error': 'Specify both branch_a and branch_b.'})
        cid_a = index['branches'].get(branch_a)
        cid_b = index['branches'].get(branch_b)
        if not cid_a:
            return jsonify({'error': f'Branch "{branch_a}" has no commits.'})
        if not cid_b:
            return jsonify({'error': f'Branch "{branch_b}" has no commits.'})
        pa = get_commit_docx(cid_a)
        pb = get_commit_docx(cid_b)
        if not pa:
            return jsonify({'error': f'File not found in branch "{branch_a}".'})
        if not pb:
            return jsonify({'error': f'File not found in branch "{branch_b}".'})
        left_paras  = extract_paragraphs(pa)
        right_paras = extract_paragraphs(pb)
        left_label  = f'{branch_a} ({cid_a[:7]})'
        right_label = f'{branch_b} ({cid_b[:7]})'

    elif mode == 'commits':
        cid_a = data.get('commit_a', '')
        cid_b = data.get('commit_b', '')
        if not cid_a or not cid_b:
            return jsonify({'error': 'Specify both commit_a and commit_b.'})
        pa = get_commit_docx(cid_a)
        pb = get_commit_docx(cid_b)
        if not pa:
            return jsonify({'error': f'File not in commit "{cid_a}".'})
        if not pb:
            return jsonify({'error': f'File not in commit "{cid_b}".'})
        left_paras  = extract_paragraphs(pa)
        right_paras = extract_paragraphs(pb)
        ca = index['commits'][cid_a]
        cb = index['commits'][cid_b]
        left_label  = f'{cid_a[:7]} — {ca["message"]}'
        right_label = f'{cid_b[:7]} — {cb["message"]}'
    else:
        return jsonify({'error': f'Unknown mode: {mode}'})

    html = build_diff_html(left_paras, right_paras, left_label, right_label)
    sid  = str(uuid.uuid4())[:8]
    DIFF_SESSIONS[sid] = html
    return jsonify({'session_id': sid, 'diff_url': f'https://127.0.0.1:5000/diff-view/{sid}'})


# ── API: command ──────────────────────────────────────────────────────────────
@app.route('/api/command', methods=['POST'])
def api_command():
    data     = request.get_json(force=True)
    command  = data.get('command', '').strip()
    filepath = normalize_path(data.get('filepath', ''))
    if not filepath:
        return jsonify({'error': 'No file path provided.'})
    doc_dir  = os.path.dirname(filepath)
    if not os.path.isdir(doc_dir):
        return jsonify({'error': f'Directory does not exist: {doc_dir}'})

    cmd_base = get_cmd_base()
    args     = shlex.split(command)
    verb     = args[0] if args else ''

    # Block switch if uncommitted changes exist
    if verb == 'switch':
        import hashlib
        idx = load_index(doc_dir)
        if idx:
            branch    = idx.get('HEAD', 'main')
            cid       = idx['branches'].get(branch)
            tree      = idx['commits'][cid]['tree'] if cid and cid in idx['commits'] else {}
            changed   = []
            try:
                working = [f for f in os.listdir(doc_dir)
                           if f.endswith('.docx') and not f.startswith('~$')
                           and os.path.isfile(os.path.join(doc_dir, f))]
                for f in working:
                    full = os.path.join(doc_dir, f)
                    h = hashlib.sha256(open(full,'rb').read()).hexdigest()
                    if f not in tree or tree[f] != h:
                        changed.append(f)
                for f in tree:
                    if f not in working:
                        changed.append(f)
            except Exception:
                pass
            if changed:
                return jsonify({
                    'output': 'BLOCKED: You have uncommitted changes in: ' +
                              ', '.join(changed) +
                              '\n\nPlease commit your changes before switching branches.',
                    'code': 1
                })

    if verb in ('commit', 'diff'):
        full_cmd = cmd_base + [verb, filepath] + args[1:]
    else:
        full_cmd = cmd_base + args

    out, code = run_subprocess(full_cmd, doc_dir)
    return jsonify({'output': out, 'code': code})


if __name__ == '__main__':
    cert = os.path.join(BASE_DIR, 'localhost+1.pem')
    key  = os.path.join(BASE_DIR, 'localhost+1-key.pem')
    if not os.path.exists(cert):
        print('ERROR: localhost+1.pem not found. Run: mkcert localhost 127.0.0.1')
        sys.exit(1)
    print('DocGit server → https://127.0.0.1:5000')
    app.run(host='127.0.0.1', port=5000, ssl_context=(cert, key), debug=False)
