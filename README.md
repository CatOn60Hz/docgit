# Docgit

Docgit is a Git-inspired version control system for Microsoft Word `.docx` files. It stores document snapshots locally, tracks branches and commits, and exposes a CLI designed around Word document workflows instead of plain text source code.

The repository also includes a simple GitHub Pages site in [index.html](/abs/path/C:/Users/Arfan/docgitt/index.html) so the project can be presented publicly from the repo itself.

## What It Does

- Initializes a local `.docgit/` repository
- Commits snapshots of `.docx` files
- Tracks branches and commit history
- Shows status and diffs for Word documents
- Switches branches and restores committed versions
- Attempts seamless Word document hot-swapping through `pywin32`

## Installation

```bash
pip install -r requirements.txt
```

Or install the package locally:

```bash
pip install .
```

## Quick Start

```bash
docgit init
docgit commit MyDocument.docx -m "Initial draft"
docgit status
docgit log
docgit diff MyDocument.docx
docgit branch edits
docgit switch edits
```

## CLI Commands

- `docgit init`
- `docgit commit <file> -m "message"`
- `docgit commit -a -m "message"`
- `docgit status`
- `docgit log`
- `docgit diff <file>`
- `docgit branch [name]`
- `docgit switch <branch>`
- `docgit checkout <branch>`
- `docgit rollback [commit]`
- `docgit merge <branch>`
- `docgit show <file>`

## Notes

- Docgit is built for `.docx` workflows, not general-purpose file versioning.
- Word document merging is intentionally conservative to avoid corrupting document structure.
- The current implementation depends on Windows-specific Word automation for hot-swapping open files.
