#!/usr/bin/env python3
"""
File Finder — a mini "Google search" for your company's files.

Two ways to use it:

  A) Search your Google Shared Drive (no downloads! just the file list):

       python3 file_finder.py scan-gdrive          <- grab file list from Drive
       python3 file_finder.py search "mpl playoff" <- search, gives Drive links

  B) Search a normal folder on your computer (reads inside documents):

       python3 file_finder.py scan /path/to/folder
       python3 file_finder.py search "onboarding checklist"

  Or run with no command for chat mode:  python3 file_finder.py
"""

import os
import sys
import json
import datetime
import argparse
import subprocess

FOLDER = os.path.dirname(os.path.abspath(__file__))
GDRIVE_INDEX_FILE = os.path.join(FOLDER, "gdrive_index.json")
LOCAL_INDEX_FILE = os.path.join(FOLDER, "index.json")

READABLE_EXTENSIONS = {".pdf", ".docx", ".xlsx", ".txt", ".md", ".csv"}

# Document types we consider "readable content" (used only for labelling)
DOC_EXTENSIONS = {".pdf", ".docx", ".doc", ".xlsx", ".xls", ".csv",
                  ".txt", ".md", ".pptx", ".ppt"}
VIDEO_EXTENSIONS = {".mp4", ".mov", ".avi", ".mkv", ".webm", ".m4v"}


def drive_link(file_id):
    """Turn a Google Drive file ID into a clickable link."""
    return f"https://drive.google.com/file/d/{file_id}/view"


def scan_gdrive():
    """Ask Google Drive for its file list (names + links only, NO downloads)."""
    print("Asking Google Drive for its file list (this downloads nothing)...\n")
    try:
        raw = subprocess.run(
            ["rclone", "lsjson", "gdrive:", "--recursive", "--files-only"],
            capture_output=True, text=True, timeout=600,
        ).stdout
    except Exception as e:
        print(f"Oops, could not talk to Google Drive: {e}")
        return

    try:
        items = json.loads(raw)
    except json.JSONDecodeError:
        print("Could not read rclone's answer. Is 'gdrive' still connected? "
              "Try:  rclone ls gdrive: --max-depth 1")
        return

    entries = []
    for item in items:
        ext = os.path.splitext(item["Name"])[1].lower()
        kind = "doc" if ext in DOC_EXTENSIONS else (
            "video" if ext in VIDEO_EXTENSIONS else "other")
        entries.append({
            "name": item["Name"],
            "path": item["Path"],
            "modified": item["ModTime"][:16].replace("T", " "),
            "size_mb": round(item["Size"] / 1024 / 1024, 1),
            "file_id": item["ID"],
            "kind": kind,
        })

    with open(GDRIVE_INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f)

    docs = sum(1 for e in entries if e["kind"] == "doc")
    vids = sum(1 for e in entries if e["kind"] == "video")
    print(f"Done! Remembered {len(entries)} files "
          f"({docs} documents, {vids} videos, {len(entries)-docs-vids} other).")
    print("Nothing was downloaded — your files stay in the cloud.\n")
    print(f"Try:  python3 file_finder.py search \"mpl playoff\"")


def search_gdrive(query, top_n=8):
    """Search the Drive file list by name. Newest versions get a highlight."""
    if not os.path.exists(GDRIVE_INDEX_FILE):
        print("No Drive list yet! First run:  python3 file_finder.py scan-gdrive")
        return

    with open(GDRIVE_INDEX_FILE, "r", encoding="utf-8") as f:
        entries = json.load(f)

    words = [w for w in query.lower().replace("_", " ").split() if w]
    if not words:
        print("Type some words to search for!")
        return

    results = []
    for entry in entries:
        name_lower = entry["name"].lower().replace("_", " ").replace("-", " ")
        path_lower = entry["path"].lower()
        score = 0
        for w in words:
            if w in name_lower:
                score += 10
            elif w in path_lower:
                score += 3
        if score > 0:
            results.append((score, entry))

    results.sort(key=lambda r: -r[0])

    if not results:        print("Sorry, nothing in the Drive matches. Try other words.")
        return

    print(f"\nFound {len(results[:top_n])} best matches (newest shown with ⭐):\n")
    shown = results[:top_n]
    newest_date = max(e["modified"] for _, e in shown)
    for i, (score, entry) in enumerate(shown, 1):
        star = " ⭐ NEWEST" if entry["modified"] == newest_date else ""
        kind_icon = "🎬" if entry["kind"] == "video" else "📄"
        print(f"{i}. {kind_icon} {entry['name']}{star}")
        print(f"   📁 {entry['path']}   (updated {entry['modified']}, "
              f"{entry['size_mb']} MB)")
        print(f"   🔗 {drive_link(entry['file_id'])}")
        print()

    if len(results) > top_n:
        print(f"(+{len(results) - top_n} more matches — try adding words to narrow it)")


def read_file_text(path):
    """Open a document on this computer and return the words inside it."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            from pypdf import PdfReader
            reader = PdfReader(path)
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        if ext == ".docx":
            import docx
            document = docx.Document(path)
            parts = [p.text for p in document.paragraphs]
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        parts.append(cell.text)
            return "\n".join(parts)
        if ext == ".xlsx":
            import openpyxl
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            parts = []
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    parts.extend(str(c) for c in row if c is not None)
            return "\n".join(parts)
        if ext in (".txt", ".md", ".csv"):
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
    except Exception:
        return ""
    return ""


def scan_folder(folder):
    """Read every document in a local folder (for the folder-search mode)."""
    if not os.path.isdir(folder):
        print(f"Oops: '{folder}' is not a folder. Check the path and try again.")
        return

    print(f"Reading all files in: {folder}\n")
    entries = []
    skipped = 0
    for root, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        for name in files:
            if name.startswith(".") or name.endswith(".partial"):
                continue
            path = os.path.join(root, name)
            ext = os.path.splitext(name)[1].lower()
            if ext not in READABLE_EXTENSIONS:
                skipped += 1
                continue
            stat = os.stat(path)
            entries.append({
                "name": name,
                "path": path,
                "modified": datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M"),
                "size_kb": round(stat.st_size / 1024, 1),
                "text": read_file_text(path),
            })
            print(f"  read: {name}")

    with open(LOCAL_INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f)

    print(f"\nDone! Remembered {len(entries)} documents "
          f"({skipped} skipped — not a readable type).")


def search_local(query, top_n=5):
    """Search local documents by name AND by the words inside them."""
    if not os.path.exists(LOCAL_INDEX_FILE):
        return False

    with open(LOCAL_INDEX_FILE, "r", encoding="utf-8") as f:
        index = json.load(f)
    if not index:
        return False

    words = [w for w in query.lower().split() if w]
    results = []
    for entry in index:
        name_lower = entry["name"].lower()
        text_lower = entry["text"].lower()
        score = 0
        for w in words:
            if w in name_lower:
                score += 10            count = text_lower.count(w)
            if count > 0:
                score += 1 + min(count, 5)
        if score > 0:
            snippet = ""
            for w in words:
                pos = text_lower.find(w)
                if pos >= 0:
                    start = max(0, pos - 40)
                    snippet = entry["text"][start:pos + 60].replace("\n", " ").strip()
                    break
            results.append((score, entry, snippet))

    if not results:
        return False

    results.sort(key=lambda r: -r[0])
    print(f"\n📄 Matches inside documents ({len(results[:top_n])}):\n")
    for i, (score, entry, snippet) in enumerate(results[:top_n], 1):
        print(f"{i}. 📄 {entry['name']}   (updated {entry['modified']})")
        print(f"   📁 {entry['path']}")
        if snippet:
            print(f"   💬 ...{snippet}...")
        print()
    return True


def search_all(query):
    """Search local documents (if scanned) and Google Drive file names."""
    search_local(query)
    print("🔎 Searching Google Drive file names...")
    search_gdrive(query)


def interactive_chat():
    print("=" * 50)
    print("  FILE FINDER — ask me where a file is!")
    print("  Type your question, or 'quit' to exit.")
    print("=" * 50)
    while True:
        try:
            query = input("\n🔎 You: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nBye!")
            return
        if query.lower() in ("quit", "exit", "q"):
            print("Bye!")
            return
        if query:
            search_all(query)


def main():
    parser = argparse.ArgumentParser(description="Find files by meaning, not just name.")
    parser.add_argument("command", nargs="?", default="chat",
                        help="scan-gdrive | scan <folder> | search <question> | chat")
    parser.add_argument("argument", nargs="?", default=None)
    args = parser.parse_args()

    if args.command == "scan-gdrive":
        scan_gdrive()
    elif args.command == "scan":
        if not args.argument:
            print("Tell me which folder, e.g.:  python3 file_finder.py scan ~/Documents")
        else:
            scan_folder(os.path.expanduser(args.argument))
    elif args.command == "search":
        if not args.argument:
            print("Tell me what to search, e.g.:  python3 file_finder.py search budget")
        else:
            search_all(args.argument)
    else:
        interactive_chat()


if __name__ == "__main__":
    main()
