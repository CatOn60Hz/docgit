import os
import json
import hashlib
import difflib
import datetime
import shutil
import uuid
from pathlib import Path
import click
from rich.console import Console
from rich.table import Table
from docx import Document

console = Console()

DOCGIT_DIR = ".docgit"
INDEX_FILE = os.path.join(DOCGIT_DIR, "index.json")
OBJECTS_DIR = os.path.join(DOCGIT_DIR, "objects")

def get_docgit_path():
    return Path(DOCGIT_DIR)

def ensure_repo():
    if not get_docgit_path().exists():
        console.print("[red]Error: Not a docgit repository. Run 'docgit init' first.[/red]")
        click.Context.exit(1)

def load_index():
    if not os.path.exists(INDEX_FILE):
        return {
            "HEAD": "main",
            "branches": {"main": None},
            "commits": {}
        }
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
        if "HEAD" not in data:
            console.print("[red]Error: Old docgit repository format detected. Because of the new branching upgrade, you need to delete the hidden '.docgit' folder in this directory and run 'docgit init' again.[/red]")
            import sys
            sys.exit(1)
        return data

def save_index(index):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=4)

def hash_file(filepath):
    hasher = hashlib.sha256()
    with open(filepath, 'rb') as f:
        buf = f.read()
        hasher.update(buf)
    return hasher.hexdigest()

def normalize_path(filepath):
    """Normalize absolute paths to relative, and ensure it's tracked correctly."""
    if filepath:
        return os.path.relpath(filepath)
    return filepath

def extract_text_from_docx(filepath):
    try:
        doc = Document(filepath)
        text_lines = []
        page_num = 1
        line_num = 1
        for p in doc.paragraphs:
            # Check for page breaks
            for run in p._element.xpath('.//w:lastRenderedPageBreak | .//w:br[@w:type="page"]'):
                page_num += 1
                
            # Check for images
            images = p._element.xpath('.//w:drawing')
            image_tags = f" [Image(s): {len(images)}]" if images else ""
            
            line_content = f"[Pg {page_num}, Ln {line_num}] {p.text}{image_tags}"
            text_lines.append(line_content)
            line_num += 1
            
        return text_lines
    except Exception as e:
        console.print(f"[red]Error reading {filepath}: {e}[/red]")
        return []

def get_current_commit(index):
    head_branch = index["HEAD"]
    commit_id = index["branches"].get(head_branch)
    if commit_id and commit_id in index["commits"]:
        return commit_id, index["commits"][commit_id]
    return None, None

def hot_swap_document(filepath, object_path):
    """
    Tries to seamlessly close the document in Microsoft Word, overwrite it, and reopen it.
    Returns True if successful, False if it couldn't connect or find the document.
    """
    try:
        import win32com.client
        word = win32com.client.GetActiveObject("Word.Application")
        abs_path = os.path.abspath(filepath)
        for doc in word.Documents:
            if doc.FullName.lower() == abs_path.lower():
                doc.Close(SaveChanges=False)
                shutil.copy2(object_path, filepath)
                word.Documents.Open(abs_path)
                return True
    except Exception:
        pass
    return False

@click.group()
def cli():
    """docgit - A Git-like version control system for Word documents."""
    pass

@cli.command()
def init():
    """Initialize a docgit repository in the current folder."""
    if get_docgit_path().exists():
        console.print("[yellow]docgit repository already exists here.[/yellow]")
        return
    
    os.makedirs(DOCGIT_DIR)
    os.makedirs(OBJECTS_DIR)
    save_index({
        "HEAD": "main",
        "branches": {"main": None},
        "commits": {}
    })
    console.print(f"[green]Initialized empty docgit repository in {os.path.abspath(DOCGIT_DIR)}[/green]")

@cli.command()
@click.argument('filepath', required=False)
@click.option('-m', '--message', required=True, help="Commit message")
@click.option('-a', '--all', 'commit_all', is_flag=True, help="Commit all modified .docx files")
def commit(filepath, message, commit_all):
    """Save a versioned snapshot. Use -a to commit all, or provide a filepath."""
    ensure_repo()
    if not filepath and not commit_all:
        console.print("[red]Error: Must specify a filepath or use -a / --all[/red]")
        return
        
    index = load_index()
    head_branch = index["HEAD"]
    curr_commit_id, curr_commit = get_current_commit(index)
    
    # Base new tree on the previous tree
    new_tree = dict(curr_commit["tree"]) if curr_commit else {}
    
    files_to_commit = []
    working_files = [f for f in os.listdir(".") if f.endswith(".docx") and not f.startswith("~$") and os.path.isfile(f)]
    
    if commit_all:
        files_to_commit = working_files
    else:
        filepath = normalize_path(filepath)
        if not os.path.exists(filepath):
            console.print(f"[red]Error: File '{filepath}' does not exist.[/red]")
            return
        if not filepath.endswith(".docx"):
            console.print("[red]Error: docgit only supports .docx files.[/red]")
            return
        files_to_commit = [filepath]

    has_changes = False
    
    for f in files_to_commit:
        f_hash = hash_file(f)
        if f not in new_tree or new_tree[f] != f_hash:
            # File changed or is new
            object_path = os.path.join(OBJECTS_DIR, f"{f_hash}.docx")
            if not os.path.exists(object_path):
                shutil.copy2(f, object_path)
            new_tree[f] = f_hash
            has_changes = True

    # Check for deleted files if we used -a
    if commit_all:
        for tracked_file in list(new_tree.keys()):
            if tracked_file not in working_files:
                del new_tree[tracked_file]
                has_changes = True

    if not has_changes:
        console.print("[yellow]No changes to commit.[/yellow]")
        return
        
    new_commit_id = str(uuid.uuid4())[:8]
    index["commits"][new_commit_id] = {
        "message": message,
        "timestamp": datetime.datetime.now().isoformat(),
        "parent": curr_commit_id,
        "tree": new_tree
    }
    
    index["branches"][head_branch] = new_commit_id
    save_index(index)
    
    console.print(f"[green][{head_branch} {new_commit_id}] {message}[/green]")

@cli.command()
def log():
    """Show commit history for the current branch."""
    ensure_repo()
    index = load_index()
    
    curr_commit_id, curr_commit = get_current_commit(index)
    
    if not curr_commit_id:
        console.print(f"[yellow]No commits on branch {index['HEAD']}.[/yellow]")
        return
        
    table = Table(title=f"Commit History (Branch: {index['HEAD']})")
    table.add_column("Commit ID", style="cyan", no_wrap=True)
    table.add_column("Date", style="magenta")
    table.add_column("Message", style="green")
    
    while curr_commit_id:
        c = index["commits"][curr_commit_id]
        table.add_row(curr_commit_id, c["timestamp"], c["message"])
        curr_commit_id = c.get("parent")
        
    console.print(table)

@cli.command()
def status():
    """Overview of tracked documents compared to the current commit."""
    ensure_repo()
    index = load_index()
    head_branch = index["HEAD"]
    
    console.print(f"On branch [bold cyan]{head_branch}[/bold cyan]")
    
    _, curr_commit = get_current_commit(index)
    tree = curr_commit["tree"] if curr_commit else {}
    
    working_files = [f for f in os.listdir(".") if f.endswith(".docx") and not f.startswith("~$") and os.path.isfile(f)]
    
    table = Table(title="docgit Status")
    table.add_column("File", style="cyan")
    table.add_column("Status", style="magenta")
    
    for f in working_files:
        if f not in tree:
            table.add_row(f, "[dim]Untracked[/dim]")
        else:
            current_hash = hash_file(f)
            if tree[f] == current_hash:
                pass # table.add_row(f, "[green]Unmodified[/green]")
            else:
                table.add_row(f, "[yellow]Modified[/yellow]")
                
    for f in tree:
        if f not in working_files:
            table.add_row(f, "[red]Deleted[/red]")
            
    if table.row_count > 0:
        console.print(table)
    else:
        console.print("Nothing to commit, working tree clean.")

@cli.command()
@click.argument('filepath')
def diff(filepath):
    """Paragraph-aware terminal diff."""
    ensure_repo()
    filepath = normalize_path(filepath)
    if not os.path.exists(filepath):
        console.print(f"[red]Error: File '{filepath}' does not exist.[/red]")
        return
        
    index = load_index()
    _, curr_commit = get_current_commit(index)
    tree = curr_commit["tree"] if curr_commit else {}
    
    if filepath not in tree:
        console.print(f"[yellow]File '{filepath}' is untracked.[/yellow]")
        return
        
    head_hash = tree[filepath]
    object_path = os.path.join(OBJECTS_DIR, f"{head_hash}.docx")
    
    if not os.path.exists(object_path):
        console.print(f"[red]Error: Snapshot missing in objects store.[/red]")
        return
        
    old_text = extract_text_from_docx(object_path)
    new_text = extract_text_from_docx(filepath)
    
    differ = difflib.SequenceMatcher(None, old_text, new_text)
    opcodes = differ.get_opcodes()
    
    has_changes = False
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "replace":
            has_changes = True
            for line in old_text[i1:i2]:
                console.print(f"[-] {line}", style="red")
            for line in new_text[j1:j2]:
                console.print(f"[+] {line}", style="green")
        elif tag == "delete":
            has_changes = True
            for line in old_text[i1:i2]:
                console.print(f"[-] {line}", style="red")
        elif tag == "insert":
            has_changes = True
            for line in new_text[j1:j2]:
                console.print(f"[+] {line}", style="green")
                
    if not has_changes:
        console.print("[yellow]No changes detected.[/yellow]")

@cli.command()
@click.argument('branch_name', required=False)
def branch(branch_name):
    """List or create branches."""
    ensure_repo()
    index = load_index()
    
    if not branch_name:
        for b in index["branches"]:
            if b == index["HEAD"]:
                console.print(f"* [green]{b}[/green]")
            else:
                console.print(f"  {b}")
        return
        
    if branch_name in index["branches"]:
        console.print(f"[red]Error: Branch '{branch_name}' already exists.[/red]")
        return
        
    curr_commit_id, _ = get_current_commit(index)
    index["branches"][branch_name] = curr_commit_id
    save_index(index)
    console.print(f"[green]Created branch '{branch_name}'[/green]")

@cli.command()
@click.argument('target')
def switch(target):
    """Switch branches automatically, auto-committing local changes."""
    ensure_repo()
    index = load_index()
    if target not in index["branches"]:
        console.print(f"[red]Error: Branch '{target}' not found.[/red]")
        return
        
    _, curr_commit = get_current_commit(index)
    tree = curr_commit["tree"] if curr_commit else {}
    working_files = [f for f in os.listdir(".") if f.endswith(".docx") and not f.startswith("~$") and os.path.isfile(f)]
    
    has_changes = False
    for f in working_files:
        if f not in tree or hash_file(f) != tree[f]:
            has_changes = True
            break
    for f in tree:
        if f not in working_files:
            has_changes = True
            break
            
    if has_changes:
        console.print("[yellow]Uncommitted changes detected. Auto-committing...[/yellow]")
        commit.callback(filepath=None, message=f"Auto-commit before switching to '{target}'", commit_all=True)
        
    checkout.callback(target)

@cli.command()
@click.argument('target')
def checkout(target):
    """Switch to a branch or restore a commit."""
    ensure_repo()
    index = load_index()
    
    if target in index["branches"]:
        commit_id = index["branches"][target]
        index["HEAD"] = target
        save_index(index)
        console.print(f"Switched to branch '{target}'")
    else:
        console.print(f"[red]Error: Branch '{target}' not found.[/red]")
        return
        
    if commit_id and commit_id in index["commits"]:
        tree = index["commits"][commit_id]["tree"]
        for filepath, f_hash in tree.items():
            object_path = os.path.join(OBJECTS_DIR, f"{f_hash}.docx")
            if os.path.exists(object_path):
                try:
                    shutil.copy2(object_path, filepath)
                except PermissionError:
                    if not hot_swap_document(filepath, object_path):
                        console.print(f"[red]Error: Cannot update '{filepath}'. Please close Microsoft Word and run the checkout command again.[/red]")
                        return
    else:
        console.print("[yellow]Branch has no commits yet.[/yellow]")

@cli.command()
@click.argument('commit_id', required=False)
def rollback(commit_id):
    """Roll back the current branch to a previous commit."""
    ensure_repo()
    index = load_index()
    head_branch = index["HEAD"]
    curr_commit_id = index["branches"].get(head_branch)
    
    if not curr_commit_id:
        console.print("[red]Error: No commits to roll back.[/red]")
        return
        
    if not commit_id:
        # Default to parent commit (1 step back)
        parent_id = index["commits"][curr_commit_id].get("parent")
        if not parent_id:
            console.print("[red]Error: Already at the first commit. Cannot roll back further.[/red]")
            return
        target_commit_id = parent_id
    else:
        # Roll back to the specific commit ID
        if commit_id not in index["commits"]:
            # Check for short ID match
            matches = [c for c in index["commits"] if c.startswith(commit_id)]
            if len(matches) == 1:
                target_commit_id = matches[0]
            elif len(matches) > 1:
                console.print(f"[red]Error: Commit ID '{commit_id}' is ambiguous.[/red]")
                return
            else:
                console.print(f"[red]Error: Commit '{commit_id}' not found.[/red]")
                return
        else:
            target_commit_id = commit_id

    # Update branch pointer
    index["branches"][head_branch] = target_commit_id
    save_index(index)
    console.print(f"[green]Rolled back branch '{head_branch}' to commit {target_commit_id[:7]}.[/green]")
    
    # Restore the working directory to match the target commit
    checkout.callback(head_branch)

def get_ancestry(index, commit_id):
    ancestors = []
    curr = commit_id
    while curr and curr in index["commits"]:
        ancestors.append(curr)
        curr = index["commits"][curr].get("parent")
    return ancestors

@cli.command()
@click.argument('branch_name')
def merge(branch_name):
    """Merge another branch into the current branch."""
    ensure_repo()
    index = load_index()
    
    if branch_name not in index["branches"]:
        console.print(f"[red]Error: Branch '{branch_name}' not found.[/red]")
        return
        
    head_branch = index["HEAD"]
    if head_branch == branch_name:
        console.print("[yellow]Cannot merge a branch into itself.[/yellow]")
        return
        
    head_commit_id = index["branches"][head_branch]
    target_commit_id = index["branches"][branch_name]
    
    if not target_commit_id:
        console.print(f"[yellow]Branch '{branch_name}' has no commits to merge.[/yellow]")
        return
        
    if not head_commit_id:
        index["branches"][head_branch] = target_commit_id
        save_index(index)
        checkout.callback(head_branch)
        console.print(f"[green]Fast-forward merged '{branch_name}' into '{head_branch}'.[/green]")
        return
        
    head_ancestors = get_ancestry(index, head_commit_id)
    if target_commit_id in head_ancestors:
        console.print(f"[yellow]Already up to date. '{branch_name}' is already merged.[/yellow]")
        return
        
    target_ancestors = get_ancestry(index, target_commit_id)
    if head_commit_id in target_ancestors:
        index["branches"][head_branch] = target_commit_id
        save_index(index)
        checkout.callback(head_branch)
        console.print(f"[green]Fast-forward merged '{branch_name}' into '{head_branch}'.[/green]")
        return
        
    lca_id = None
    for ancestor in head_ancestors:
        if ancestor in target_ancestors:
            lca_id = ancestor
            break
            
    lca_tree = index["commits"][lca_id]["tree"] if lca_id else {}
    head_tree = index["commits"][head_commit_id]["tree"]
    target_tree = index["commits"][target_commit_id]["tree"]
    
    new_tree = dict(head_tree)
    
    for f in set(list(head_tree.keys()) + list(target_tree.keys())):
        head_hash = head_tree.get(f)
        target_hash = target_tree.get(f)
        lca_hash = lca_tree.get(f)
        
        if head_hash != lca_hash and target_hash != lca_hash:
            if head_hash != target_hash:
                console.print(f"[red]Error: Merge conflict in '{f}'. Both branches modified this file.[/red]")
                console.print("[red]Docgit cannot safely merge Word documents automatically. Please manually resolve this before merging.[/red]")
                return
        
        if target_hash != lca_hash and head_hash == lca_hash:
            new_tree[f] = target_hash
            
    commit_id = hashlib.sha1(str(time.time()).encode()).hexdigest()[:8]
    index["commits"][commit_id] = {
        "id": commit_id,
        "message": f"Merge branch '{branch_name}' into '{head_branch}'",
        "timestamp": datetime.datetime.now().isoformat(),
        "parent": head_commit_id,
        "merge_parent": target_commit_id,
        "tree": new_tree
    }
    index["branches"][head_branch] = commit_id
    save_index(index)
    
    checkout.callback(head_branch)
    console.print(f"[green]Successfully merged '{branch_name}' into '{head_branch}'.[/green]")

@cli.command()
@click.argument('filepath')
def show(filepath):
    """Print latest committed text for a file."""
    ensure_repo()
    filepath = normalize_path(filepath)
    index = load_index()
    _, curr_commit = get_current_commit(index)
    
    if not curr_commit or filepath not in curr_commit["tree"]:
        console.print(f"[red]Error: File '{filepath}' is untracked in current branch.[/red]")
        return
        
    head_hash = curr_commit["tree"][filepath]
    object_path = os.path.join(OBJECTS_DIR, f"{head_hash}.docx")
    
    if not os.path.exists(object_path):
        console.print(f"[red]Error: Snapshot missing in objects store.[/red]")
        return
        
    text_lines = extract_text_from_docx(object_path)
    for line in text_lines:
        print(line)

if __name__ == '__main__':
    cli()
