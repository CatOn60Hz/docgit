# How Docgit Works: Architecture & Internal Mechanics

Docgit is a custom, Git-inspired version control system specifically engineered for Microsoft Word (`.docx`) documents. Unlike traditional Git (which is optimized for plain text code), Docgit is designed to handle binary zip-archives (which `.docx` files actually are), while providing semantic text-level diffing and deep integration with the Microsoft Office ecosystem.

## 1. The Local Repository (`.docgit/`)

When you run `docgit init`, a hidden folder named `.docgit` is created in your directory. This serves as the local database for your version history.

It contains two main components:
*   **`objects/` directory:** The content-addressable storage. Every time you commit, Docgit hashes the raw binary contents of the `.docx` file using `SHA-256`. The document is then copied into the `objects/` folder and named by its hash (e.g., `a4c9b...8d.docx`). If you commit the exact same document twice, it doesn't waste space because the hash will be identical.
*   **`index.json`:** The brain of the repository. It tracks the commit graph, the branches, and the current state of the workspace.

## 2. Commits & The "Tree" System

A commit in Docgit is a snapshot of the entire directory of tracked `.docx` files at a specific point in time. 
Inside `index.json`, a commit looks like this:
```json
"commits": {
  "ab12c34d": {
    "message": "First draft of conclusion",
    "timestamp": "2026-05-05T10:00:00",
    "parent": "f89e21a0",
    "tree": {
      "Research_Paper.docx": "a4c9b...8d",
      "References.docx": "b7d2f...1e"
    }
  }
}
```
The **`tree`** maps your human-readable file paths (`Research_Paper.docx`) to the immutable binary snapshots stored in the `objects/` folder.

## 3. Branching & The HEAD Pointer

Branches are simply lightweight text pointers to specific commit IDs. 
```json
"branches": {
  "main": "ab12c34d",
  "experimental": "c98f12a3"
},
"HEAD": "experimental"
```
*   **`HEAD`** tells Docgit which branch you currently have checked out in your working directory.
*   When you run `docgit switch main`, Docgit reads `main`'s commit ID, looks up the `tree`, and copies those exact binary files from the `objects/` folder back into your main working directory.

## 4. Seamless Microsoft Word Integration ("Hot-Swapping")

The biggest challenge with version-controlling Word documents is that Microsoft Word tightly **locks** files while they are open. If Docgit tries to overwrite a file during a `checkout` or `rollback` while Word is open, the OS will throw a `PermissionError`.

To solve this, Docgit uses the `pywin32` library to hook directly into the active Windows COM (Component Object Model) interface of Microsoft Word. 
When `docgit` detects a lock, it:
1. Connects to the active Word instance.
2. Finds the specific document.
3. Closes it invisibly without saving.
4. Overwrites the file on the hard drive with the new branch's version.
5. Instantly reopens the document in Word.

This allows users to switch branches from the CLI or Ribbon without ever leaving the Word interface.

## 5. Diffing Engine & XML Parsing

Because `.docx` files are binary zip archives, you cannot diff them directly. When you run `docgit diff`, Docgit uses the `python-docx` library to extract the text.

However, Docgit goes deeper than standard text extraction:
*   It parses the underlying XML (`document.xml`) of the Word document.
*   It looks for explicit and rendered page break tags (`<w:lastRenderedPageBreak>`) to calculate **Page Numbers**.
*   It counts paragraphs to calculate **Line Numbers**.
*   It detects `Drawing` objects (`<w:drawing>`) to flag when **Images** are added or removed.

This metadata is prefixed to the text (e.g., `[Pg 2, Ln 4] [Image(s): 1] Hello World`), which is then fed into Python's `difflib.unified_diff` to generate a colorized, human-readable terminal diff.

## 6. Merging Logic

Docgit supports two types of merges via `docgit merge <branch>`:
1.  **Fast-Forward:** If the target branch is a direct descendant of the current branch, Docgit simply moves the branch pointer forward.
2.  **3-Way File-Level Merge:** If both branches have diverged, Docgit finds the **Lowest Common Ancestor (LCA)**. It then compares the files in both branches against the LCA. 
    *   If Branch A modified `File1` and Branch B modified `File2`, Docgit safely combines them into a single new merge commit.
    *   **Conflict Detection:** If Docgit detects that **both** branches modified `File1`, it safely aborts the merge. Unlike code (which can be merged line-by-line), Word document XML structures are highly fragile and prone to corruption if merged textually. Aborting protects the document's formatting.

## 7. Distribution (PyInstaller & VBA)

For distribution, the Python system is bundled into a single standalone `docgit.exe` using `PyInstaller`.
The user interacts with a **VBA Macro Ribbon** embedded in a Word Template (`.dotm`). When a user clicks "Commit" on the ribbon, the VBA macro uses the Windows shell to execute the hidden `docgit.exe` in the background, effectively bridging the gap between Microsoft's UI and the Python CLI engine.
