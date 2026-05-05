import os
import json
import hashlib
import difflib
import datetime
import shutil
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
        return {"files": {}}
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_index(index):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=4)

def hash_file(filepath):
    """Generate SHA-256 hash of a file."""
    hasher = hashlib.sha256()
    with open(filepath, 'rb') as f:
        buf = f.read()
        hasher.update(buf)
    return hasher.hexdigest()

def extract_text_from_docx(filepath):
    """Extract paragraphs text from a docx file."""
    try:
        doc = Document(filepath)
        return [p.text for p in doc.paragraphs]
    except Exception as e:
        console.print(f"[red]Error reading {filepath}: {e}[/red]")
        return []

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
    save_index({"files": {}})
    console.print(f"[green]Initialized empty docgit repository in {os.path.abspath(DOCGIT_DIR)}[/green]")

@cli.command()
@click.argument('filepath')
@click.option('-m', '--message', required=True, help="Commit message")
def commit(filepath, message):
    """Save a versioned snapshot of a .docx file."""
    ensure_repo()
    if not os.path.exists(filepath):
        console.print(f"[red]Error: File '{filepath}' does not exist.[/red]")
        return
    
    if not filepath.endswith(".docx"):
        console.print("[red]Error: docgit only supports .docx files.[/red]")
        return

    # Compute hash of the file
    file_hash = hash_file(filepath)
    object_path = os.path.join(OBJECTS_DIR, f"{file_hash}.docx")
    
    index = load_index()
    if filepath not in index["files"]:
        index["files"][filepath] = {"commits": [], "head": None}
    
    file_record = index["files"][filepath]
    
    # Check if anything changed
    if file_record["head"] == file_hash:
        console.print("[yellow]No changes to commit. Working tree clean.[/yellow]")
        return
    
    # Store snapshot
    shutil.copy2(filepath, object_path)
    
    # Create commit record
    timestamp = datetime.datetime.now().isoformat()
    commit_id = file_hash[:8] # Short hash
    commit_record = {
        "id": commit_id,
        "hash": file_hash,
        "message": message,
        "timestamp": timestamp
    }
    
    file_record["commits"].append(commit_record)
    file_record["head"] = file_hash
    save_index(index)
    
    console.print(f"[green]Committed {filepath} (id: {commit_id}): {message}[/green]")

@cli.command()
@click.argument('filepath')
def log(filepath):
    """Show commit history of a file."""
    ensure_repo()
    index = load_index()
    
    if filepath not in index["files"]:
        console.print(f"[red]Error: File '{filepath}' is not tracked.[/red]")
        return
        
    file_record = index["files"][filepath]
    commits = file_record["commits"]
    
    if not commits:
        console.print(f"[yellow]No commits for {filepath}.[/yellow]")
        return
        
    table = Table(title=f"Commit History for {filepath}")
    table.add_column("Commit ID", style="cyan", no_wrap=True)
    table.add_column("Date", style="magenta")
    table.add_column("Message", style="green")
    
    for commit in reversed(commits):
        table.add_row(commit["id"], commit["timestamp"], commit["message"])
        
    console.print(table)

@cli.command()
@click.argument('filepath')
def diff(filepath):
    """Paragraph-aware terminal diff."""
    ensure_repo()
    if not os.path.exists(filepath):
        console.print(f"[red]Error: File '{filepath}' does not exist.[/red]")
        return
        
    index = load_index()
    if filepath not in index["files"] or not index["files"][filepath]["head"]:
        console.print(f"[yellow]File '{filepath}' is untracked or has no commits.[/yellow]")
        return
        
    file_record = index["files"][filepath]
    head_hash = file_record["head"]
    object_path = os.path.join(OBJECTS_DIR, f"{head_hash}.docx")
    
    if not os.path.exists(object_path):
        console.print(f"[red]Error: Snapshot for head missing in objects store.[/red]")
        return
        
    old_text = extract_text_from_docx(object_path)
    new_text = extract_text_from_docx(filepath)
    
    differ = difflib.SequenceMatcher(None, old_text, new_text)
    opcodes = differ.get_opcodes()
    
    has_changes = False
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            # Just print a context line if needed, or skip to save space
            pass
        elif tag == "replace":
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
def status():
    """Overview of all tracked documents."""
    ensure_repo()
    index = load_index()
    
    # Find all .docx files in current dir
    working_files = [f for f in os.listdir(".") if f.endswith(".docx") and os.path.isfile(f)]
    tracked_files = index["files"]
    
    table = Table(title="docgit Status")
    table.add_column("File", style="cyan")
    table.add_column("Status", style="magenta")
    
    for f in working_files:
        if f not in tracked_files:
            table.add_row(f, "[dim]Untracked[/dim]")
        else:
            file_record = tracked_files[f]
            head_hash = file_record["head"]
            current_hash = hash_file(f)
            if head_hash == current_hash:
                table.add_row(f, "[green]Unmodified[/green]")
            else:
                table.add_row(f, "[yellow]Modified[/yellow]")
                
    # Check for deleted files
    for f in tracked_files:
        if f not in working_files:
            table.add_row(f, "[red]Deleted from working directory[/red]")
            
    console.print(table)

@cli.command()
@click.argument('filepath')
@click.argument('commit_id')
def checkout(filepath, commit_id):
    """Restore a previous version."""
    ensure_repo()
    index = load_index()
    
    if filepath not in index["files"]:
        console.print(f"[red]Error: File '{filepath}' is not tracked.[/red]")
        return
        
    file_record = index["files"][filepath]
    commits = file_record["commits"]
    
    target_commit = next((c for c in commits if c["id"] == commit_id or c["hash"] == commit_id), None)
    
    if not target_commit:
        console.print(f"[red]Error: Commit '{commit_id}' not found for {filepath}.[/red]")
        return
        
    object_path = os.path.join(OBJECTS_DIR, f"{target_commit['hash']}.docx")
    if not os.path.exists(object_path):
        console.print(f"[red]Error: Snapshot missing for this commit.[/red]")
        return
        
    shutil.copy2(object_path, filepath)
    console.print(f"[green]Restored {filepath} to commit {target_commit['id']}[/green]")

@cli.command()
@click.argument('filepath')
def show(filepath):
    """Print latest committed text."""
    ensure_repo()
    index = load_index()
    
    if filepath not in index["files"] or not index["files"][filepath]["head"]:
        console.print(f"[red]Error: File '{filepath}' is untracked or has no commits.[/red]")
        return
        
    file_record = index["files"][filepath]
    head_hash = file_record["head"]
    object_path = os.path.join(OBJECTS_DIR, f"{head_hash}.docx")
    
    if not os.path.exists(object_path):
        console.print(f"[red]Error: Snapshot for head missing in objects store.[/red]")
        return
        
    text_lines = extract_text_from_docx(object_path)
    for line in text_lines:
        print(line)

if __name__ == '__main__':
    cli()
