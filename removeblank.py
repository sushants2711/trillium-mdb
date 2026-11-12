#!/usr/bin/env python3
"""
Remove Blank Pages: one interactive script, not an application or service.

QUICK START (Windows, Python 3.10+)
    py -m pip install "PyMuPDF>=1.24,<2" pywin32
    py removeblank.py

PDF works on Windows/macOS/Linux with PyMuPDF. Word (.docx/.doc) requires
Windows and installed desktop Microsoft Word, plus pywin32. No web app,
database, cloud API registration, or API key is needed.

WORKFLOW
    Choose local or OneDrive/SharePoint synced storage -> give file path ->
    count blank pages -> delete all / select pages / skip -> verify ->
    choose output folder -> save original-name_rm_whitepage.ext -> read back.

SAFETY AND SCOPE
* The source is never opened for writing. Work happens on a temporary copy.
* "Blank" means strictly white at 144 and 288 DPI, with no extracted text,
  vector drawing, link, annotation, or form field. Headers/page numbers count
  as content. Noisy scans and near-white pages are deliberately kept.
* Detection is conservative, not a guarantee about every possible PDF object.
  Review candidate page numbers before approving deletion.
* Every kept page must have the same rendered pixels, extracted text, page
  geometry, annotation data, and form data before and after deletion.
* Word: only simple empty paragraphs/manual-page-break pages are eligible.
  Section breaks, tables, anchored objects, protected documents, revisions,
  comments, signatures, and complex layout cases are not forcibly changed.
  Automatic renumbering/reflow can fail strict verification; in that case no
  output is saved. Export to PDF first if fixed-page removal is preferred.
* PDFs with layers or signature fields, and signed Word files, are refused.
  Encrypted PDFs require their owner password; encryption is retained.
* Existing files are never overwritten. If needed, a numeric suffix is added.
* OneDrive/SharePoint support is via an ALREADY SYNCED File Explorer folder.
  The script asks permission and can launch OneDrive setup. Microsoft handles
  sign-in. It does not grant SharePoint permissions, accept web URLs, or perform
  direct Graph/OAuth uploads. Local write access does not prove cloud access.
  Cloud upload completion must be checked in OneDrive / SharePoint.
* Editing may affect links/bookmarks whose destinations are removed pages.
  This is not a semantic validator for PDF accessibility tags or Word fields.
* Use only on documents you trust. Closing a preview helps temporary cleanup.

REFERENCE DOCUMENTATION
https://pymupdf.readthedocs.io/en/latest/document.html
https://pymupdf.readthedocs.io/en/latest/tutorial.html
https://learn.microsoft.com/en-us/office/vba/api/word.range.goto
https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat
https://support.microsoft.com/en-us/onedrive/sync/documents-and-library/view-sharepoint-files-in-file-explorer
"""

from __future__ import annotations

import hashlib
import math
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from getpass import getpass
import uuid
import zipfile

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz  # PyMuPDF's older import name.
    except ImportError:
        fitz = None

DPI = 144
MAX_PIXELS = 60_000_000
ALLOWED = {".pdf", ".docx", ".doc"}


class SafetyStop(RuntimeError):
    """A failed safety condition: do not publish a result."""


class Cancelled(RuntimeError):
    """The user chose not to continue."""


@dataclass(frozen=True)
class PageState:
    blank: bool
    geometry: tuple
    render_hash: str
    text: str
    objects: tuple


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def require_dependencies(word: bool = False) -> None:
    if fitz is None or not hasattr(fitz, "open"):
        raise SafetyStop('Install PyMuPDF first: py -m pip install "PyMuPDF>=1.24,<2"')
    if word and os.name != "nt":
        raise SafetyStop(
            "Word processing requires Windows and desktop Microsoft Word. "
            "On this system, export your Word file to PDF first."
        )


def render(page, dpi: int):
    pixels = math.ceil(page.rect.width * dpi / 72) * math.ceil(
        page.rect.height * dpi / 72
    )
    if pixels > MAX_PIXELS:
        raise SafetyStop(
            f"Page {page.number + 1} is too large for the safe rendering limit."
        )
    return page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB, alpha=False, annots=True)


def page_state(page) -> PageState:
    text = page.get_text("text")
    annotations = tuple(
        (tuple(a.type), tuple(a.rect), repr(a.info), repr(a.colors), a.flags)
        for a in (page.annots() or [])
    )
    widgets = tuple(
        (w.field_name, w.field_type, str(w.field_value), tuple(w.rect))
        for w in (page.widgets() or [])
    )
    # Even invisible text, off-page content, annotations, and links are protected.
    protected = bool(
        text.strip()
        or annotations
        or widgets
        or page.get_links()
        or page.get_drawings()
    )
    if not protected:
        # Check whole embedded images too: a white crop must not conceal
        # nonwhite content outside the visible page. Inline/unknown images
        # are kept rather than guessed at.
        image_info = page.get_image_info(xrefs=True)
        if any(not info.get("xref") for info in image_info):
            protected = True
        for entry in page.get_images(full=True):
            xref = entry[0]
            if not xref or entry[2] * entry[3] > MAX_PIXELS:
                protected = True
                break
            image = fitz.Pixmap(page.parent, xref)
            if image.colorspace is None:
                protected = True
                break
            if image.colorspace.n != 3:
                image = fitz.Pixmap(fitz.csRGB, image)
            if image.alpha:
                image = fitz.Pixmap(image, 0)
            image_samples = image.samples
            if image_samples.count(255) != len(image_samples):
                protected = True
                break
    pix = render(page, DPI)
    samples = pix.samples
    white = samples.count(255) == len(samples)
    blank = not protected and white
    if blank:
        # An independent higher-resolution pass can catch small raster marks.
        high = render(page, DPI * 2).samples
        blank = high.count(255) == len(high)
    geometry = (
        tuple(page.mediabox),
        tuple(page.cropbox),
        tuple(page.rect),
        page.rotation,
        pix.width,
        pix.height,
    )
    return PageState(
        blank,
        geometry,
        hashlib.sha256(samples).hexdigest(),
        text,
        (annotations, widgets),
    )


def open_pdf(path: Path, password: str = ""):
    doc = fitz.open(str(path))
    try:
        if not doc.is_pdf:
            raise SafetyStop("The file is not a valid PDF.")
        if doc.needs_pass and not doc.authenticate(password):
            raise SafetyStop("The PDF password is missing or incorrect.")
        return doc
    except BaseException:
        doc.close()
        raise


def pdf_encrypted(doc) -> bool:
    return doc.xref_get_key(-1, "Encrypt")[0] != "null"


def pdf_preflight(path: Path, password: str = "") -> None:
    with open_pdf(path, password) as doc:
        if not doc.page_count:
            raise SafetyStop("The PDF contains no pages.")
        if doc.is_repaired:
            raise SafetyStop("This PDF required repair. Repair a separate copy first.")
        if doc.get_ocgs():
            raise SafetyStop(
                "PDF layers can hide content. This script will not delete its pages."
            )
        if doc.get_sigflags() > 0 or any(
            w.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE
            for p in doc
            for w in (p.widgets() or [])
        ):
            raise SafetyStop(
                "The PDF has a signature field. Editing could invalidate a signature."
            )
        # Require owner authentication, not just a password that permits viewing.
        if pdf_encrypted(doc) and not (doc.authenticate(password) & 4):
            raise SafetyStop("An encrypted PDF requires the owner password.")
        if not (doc.permissions & fitz.PDF_PERM_MODIFY):
            raise SafetyStop("The PDF does not allow document modification.")


def scan_pdf(path: Path, password: str = "", progress: bool = True) -> list[PageState]:
    states = []
    with open_pdf(path, password) as doc:
        for i, page in enumerate(doc):
            if progress:
                print(
                    f"\rChecking page {i + 1}/{doc.page_count} ...", end="", flush=True
                )
            states.append(page_state(page))
    if progress:
        print()
    return states


def verify_pdf(
    baseline: list[PageState],
    candidate: Path,
    deleted: set[int],
    password: str = "",
    progress: bool = True,
) -> list[PageState]:
    if not deleted or not deleted.issubset(set(range(1, len(baseline) + 1))):
        raise SafetyStop("The deletion selection is empty or out of range.")
    if any(not baseline[n - 1].blank for n in deleted):
        raise SafetyStop("A selected page is not classified as blank.")
    expected = [(n, p) for n, p in enumerate(baseline, 1) if n not in deleted]
    if not expected:
        raise SafetyStop("All pages are blank; an empty document will not be created.")
    actual = scan_pdf(candidate, password, progress)
    if len(actual) != len(expected):
        raise SafetyStop(
            f"Verification failed: expected {len(expected)} pages, found {len(actual)}."
        )
    for new_n, ((old_n, before), after) in enumerate(zip(expected, actual), 1):
        if before != after:
            raise SafetyStop(
                f"Verification failed: original page {old_n} changed "
                f"at output page {new_n}. No result will be published."
            )
    return actual


def delete_pdf(
    source: Path,
    target: Path,
    baseline: list[PageState],
    deleted: set[int],
    password: str = "",
) -> list[PageState]:
    if target.exists() or target.resolve() == source.resolve():
        raise SafetyStop("The temporary output must be a new file.")
    if not deleted or any(
        n < 1 or n > len(baseline) or not baseline[n - 1].blank for n in deleted
    ):
        raise SafetyStop("Only detected blank pages may be deleted.")
    if len(deleted) == len(baseline):
        raise SafetyStop("Cannot remove every page. Original left unchanged.")
    with open_pdf(source, password) as doc:
        doc.delete_pages(sorted(n - 1 for n in deleted))
        doc.save(
            str(target),
            garbage=3,
            deflate=True,
            encryption=fitz.PDF_ENCRYPT_KEEP,
        )
    return verify_pdf(baseline, target, deleted, password)


@dataclass(frozen=True)
class WordEdit:
    start: int
    end: int


class WordSession:
    """A separate Word instance; only a disposable copy is opened for editing."""

    def __init__(self):
        self.app = None
        self.doc = None
        self.saved_options = {}
        self.pythoncom = None

    def __enter__(self):
        require_dependencies(word=True)
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise SafetyStop("Install Word support: py -m pip install pywin32") from exc
        self.pythoncom = pythoncom
        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.DispatchEx("Word.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = 0
            self.app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            for key in (
                "UpdateLinksAtOpen",
                "UpdateFieldsAtPrint",
                "UpdateLinksAtPrint",
                "BackgroundSave",
            ):
                self.saved_options[key] = getattr(self.app.Options, key)
                setattr(self.app.Options, key, False)
            return self
        except BaseException:
            self.__exit__(None, None, None)
            raise

    def close_document(self):
        if self.doc is not None:
            self.doc.Close(SaveChanges=0)
            self.doc = None

    def __exit__(self, *unused):
        try:
            self.close_document()
        finally:
            if self.app is not None:
                for key, value in self.saved_options.items():
                    try:
                        setattr(self.app.Options, key, value)
                    except Exception:
                        pass
                try:
                    self.app.Quit(SaveChanges=0)
                except Exception:
                    pass
                self.app = None
            if self.pythoncom is not None:
                self.pythoncom.CoUninitialize()
                self.pythoncom = None

    def open(self, path: Path, read_only: bool = False):
        self.close_document()
        # Supplying a random password avoids an invisible password dialog.
        # Password-protected Word documents are deliberately unsupported.
        sentinel = uuid.uuid4().hex
        self.doc = self.app.Documents.Open(
            FileName=str(path),
            ConfirmConversions=False,
            ReadOnly=read_only,
            AddToRecentFiles=False,
            Visible=False,
            PasswordDocument=sentinel,
            PasswordTemplate=sentinel,
            WritePasswordDocument=sentinel,
            WritePasswordTemplate=sentinel,
            Revert=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
        if self.doc.ProtectionType != -1:
            raise SafetyStop("Protected Word documents are not modified.")
        if self.doc.Signatures.Count:
            raise SafetyStop("Signed Word documents are not modified.")
        if self.doc.TrackRevisions or self.doc.Revisions.Count:
            raise SafetyStop(
                "Word tracked changes are present. Resolve them in a separate copy first."
            )
        if self.doc.Comments.Count:
            raise SafetyStop(
                "Word comments are present; this conservative mode will stop."
            )
        if self.doc.ReadOnly and not read_only:
            raise SafetyStop("Word opened the working copy read-only.")
        self.doc.Repaginate()

    def export_pdf(self, target: Path):
        self.doc.Repaginate()
        self.doc.ExportAsFixedFormat(
            OutputFileName=str(target),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            Item=0,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )

    def page_range(self, number: int, count: int):
        start = self.doc.GoTo(What=1, Which=1, Count=number)
        # Information(3) is the physical page, not the adjusted section number.
        if start.Information(3) != number:
            raise SafetyStop("This is an automatically inserted section/layout page.")
        end = (
            self.doc.GoTo(What=1, Which=1, Count=number + 1).Start
            if number < count
            else self.doc.Content.End - 1
        )
        return self.doc.Range(start.Start, end)

    def plan(self, number: int, count: int) -> WordEdit:
        rng = self.page_range(number, count)
        start, end = rng.Start, rng.End
        # On the last page, the manual break immediately before it may need
        # removal too. Section-break positions are checked below.
        if number == count and start > 0:
            previous = self.doc.Range(start - 1, start).Text
            if previous == "\x0c":
                start -= 1
        if start >= end:
            raise SafetyStop(
                "A required final paragraph or layout rule creates this page."
            )
        rng = self.doc.Range(start, end)
        if any(c not in " \t\r\n\x0b\x0c" for c in (rng.Text or "")):
            raise SafetyStop("The page contains hidden text or structured content.")
        if rng.Tables.Count or rng.InlineShapes.Count or rng.ContentControls.Count:
            raise SafetyStop("The page contains a table, object, or content control.")
        if rng.Fields.Count or rng.Bookmarks.Count:
            raise SafetyStop("The page contains a field or bookmark.")
        for index, section in enumerate(self.doc.Sections, 1):
            if index < self.doc.Sections.Count:
                if start <= section.Range.End - 1 < end:
                    raise SafetyStop("Deleting this page would remove a section break.")
        for shape in self.doc.Shapes:
            if start <= shape.Anchor.Start < end:
                raise SafetyStop("An object is anchored to this page.")
        return WordEdit(start, end)

    def apply(self, plans: list[WordEdit]):
        # Original offsets remain valid when deletion proceeds from the end.
        # Adjacent selected blank pages may share a trailing manual page break.
        merged = []
        for plan in sorted(plans, key=lambda p: p.start):
            if merged and plan.start <= merged[-1].end:
                last = merged.pop()
                merged.append(WordEdit(last.start, max(last.end, plan.end)))
            else:
                merged.append(plan)
        for plan in reversed(merged):
            self.doc.Range(plan.start, plan.end).Delete()
        self.doc.Repaginate()
        self.doc.Save()


def word_preflight(path: Path):
    if path.suffix.lower() == ".docx":
        if not zipfile.is_zipfile(path):
            raise SafetyStop(
                "This .docx is encrypted or not standard DOCX. "
                "Save an unprotected copy in Word first."
            )
        with zipfile.ZipFile(path) as package:
            names = package.namelist()
            if any(n.lower().startswith("_xmlsignatures/") for n in names):
                raise SafetyStop("Digitally signed Word documents are not modified.")
            if any(n.lower().endswith("vbaproject.bin") for n in names):
                raise SafetyStop("Macro-enabled Word files are not supported.")


def ask(prompt: str) -> str:
    value = input(prompt).strip()
    if value.lower() in {"quit", "exit", "cancel"}:
        raise Cancelled()
    return value


def yes(prompt: str) -> bool:
    return ask(prompt + " [y/N]: ").lower() in {"y", "yes"}


def choose(prompt: str, allowed: set[str], default: str) -> str:
    while True:
        choice = ask(prompt).lower() or default
        if choice in allowed:
            return choice
        print("Choose " + ", ".join(sorted(allowed)) + ", or type cancel.")


def open_viewer(path: Path):
    if os.name == "nt":
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def grant_synced_access(purpose: str):
    print(f"\nOneDrive / SharePoint access for {purpose}")
    print("This uses a synced File Explorer folder, not a website URL.")
    if not yes("Allow this script to access the synced folder you select"):
        raise Cancelled()
    for key in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        if os.environ.get(key):
            print(f"  {key}: {os.environ[key]}")
    print(
        "Sign in with the OneDrive app. For SharePoint, sign in to your library\n"
        "and use 'Add shortcut to OneDrive' or 'Sync'. Your organization must\n"
        "allow syncing, and you need write permission to upload an output.\n"
        "This script cannot grant missing permissions or verify remote access."
    )
    if not yes("Are you signed in and is the required folder visible in File Explorer"):
        if os.name == "nt" and yes("Open the installed OneDrive app for sign-in"):
            roots = [
                os.environ.get(k)
                for k in ("LOCALAPPDATA", "ProgramFiles", "ProgramFiles(x86)")
            ]
            executable = next(
                (
                    Path(root) / "Microsoft" / "OneDrive" / "OneDrive.exe"
                    for root in roots
                    if root
                    and (
                        Path(root) / "Microsoft" / "OneDrive" / "OneDrive.exe"
                    ).is_file()
                ),
                None,
            )
            if executable:
                subprocess.Popen([str(executable)])
            else:
                print("OneDrive was not found. Open it from the Windows Start menu.")
        ask("Complete sign-in/sync outside this script, then press Enter (or cancel): ")
        if not yes("Is the synced folder now accessible"):
            raise Cancelled()


def pick_path(directory: bool) -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        try:
            return (
                filedialog.askdirectory(title="Choose output folder")
                if directory
                else filedialog.askopenfilename(
                    title="Choose a PDF or Word document",
                    filetypes=[
                        ("Documents", "*.pdf *.docx *.doc"),
                        ("All files", "*.*"),
                    ],
                )
            )
        finally:
            root.destroy()
    except Exception as exc:
        print(f"File picker unavailable ({exc}). Paste the path instead.")
        return ""


def input_path(
    prompt: str, directory: bool = False, default: Path | None = None
) -> Path:
    while True:
        value = ask(prompt)
        if value.lower() == "browse":
            value = pick_path(directory)
            if not value:
                continue
        if not value and default is not None:
            path = default
        else:
            value = value.strip("\"'")
            if value.lower().startswith(("http://", "https://")):
                print("Use the synced local path from File Explorer, not the web URL.")
                continue
            if not value:
                continue
            path = Path(os.path.expandvars(value)).expanduser()
        try:
            path = path.resolve(strict=True)
            if directory and not path.is_dir():
                raise ValueError("Choose an existing folder.")
            if not directory and not path.is_file():
                raise ValueError("Choose an existing file.")
            if not directory and path.suffix.lower() not in ALLOWED:
                raise ValueError("Supported types: .pdf, .docx, .doc.")
            return path
        except (OSError, ValueError, RuntimeError) as exc:
            print(f"Cannot use that location: {exc}")


def parse_pages(value: str, eligible: set[int]) -> set[int]:
    selected = set()
    for item in value.split(","):
        item = item.strip()
        if not item:
            raise ValueError("Enter page numbers such as 2,5,8-10.")
        parts = item.split("-")
        if len(parts) == 1:
            selected.add(int(parts[0]))
        elif len(parts) == 2:
            first, last = map(int, parts)
            if first > last or last - first > len(eligible):
                raise ValueError("Invalid page range.")
            selected.update(range(first, last + 1))
        else:
            raise ValueError("Invalid page range.")
    if not selected or not selected.issubset(eligible):
        raise ValueError("Select only page numbers from the eligible blank-page list.")
    return selected


def select_pages(
    baseline: list[PageState], eligible: set[int], preview: Path
) -> set[int]:
    blank = {n for n, p in enumerate(baseline, 1) if p.blank}
    print(f"\nTotal pages: {len(baseline)}")
    print(f"Strictly blank pages found: {len(blank)}")
    print("Blank page numbers: " + (", ".join(map(str, sorted(blank))) or "none"))
    print(
        "Eligible for safe removal: "
        + (", ".join(map(str, sorted(eligible))) or "none")
    )
    if not eligible:
        print("Nothing eligible to delete. Original unchanged.")
        return set()
    if len(blank) == len(baseline):
        print("Every page is blank. The script will not create an empty document.")
        return set()
    print(
        "Review these page numbers. Headers, footers and page numbers count as content."
    )
    while True:
        action = choose(
            "\n[D]elete eligible pages / [S]kip / [P]ick page numbers / [V]iew PDF "
            "(default S): ",
            {"d", "s", "p", "v"},
            "s",
        )
        if action == "s":
            return set()
        if action == "v":
            try:
                open_viewer(preview)
                print(
                    "The full PDF preview is open. Close it when you finish reviewing."
                )
            except OSError as exc:
                print(f"Could not open the preview: {exc}")
            continue
        if action == "d":
            selected = eligible.copy()
        else:
            try:
                selected = parse_pages(
                    ask("Original page numbers (e.g. 2,5-7): "), eligible
                )
            except ValueError as exc:
                print(exc)
                continue
        print("Selected: " + ", ".join(map(str, sorted(selected))))
        if yes(f"Remove these {len(selected)} page(s) from a NEW copy"):
            return selected


def check_writable(folder: Path):
    # This is a local-filesystem check, not a test of SharePoint upload permission.
    with tempfile.TemporaryFile(dir=str(folder)) as stream:
        stream.write(b"blank-page-remover write test")
        stream.flush()
        os.fsync(stream.fileno())


def publish_verified_copy(
    verified: Path, folder: Path, source: Path, original_hash: str
) -> Path:
    if file_hash(source) != original_hash:
        raise SafetyStop("The original changed during processing. Start again.")
    check_writable(folder)
    verified_hash = file_hash(verified)
    stem = source.stem + "_rm_whitepage"
    suffix = source.suffix
    for number in range(10_000):
        extra = "" if number == 0 else f"_{number}"
        output = folder / f"{stem}{extra}{suffix}"
        if output.resolve() == source.resolve():
            continue
        try:
            dest = output.open("xb")  # Exclusive create: never overwrite anything.
        except FileExistsError:
            continue
        try:
            with dest, verified.open("rb") as src:
                shutil.copyfileobj(src, dest, length=1024 * 1024)
                dest.flush()
                os.fsync(dest.fileno())
            if file_hash(output) != verified_hash:
                raise SafetyStop("Saved-file read-back verification failed.")
            return output
        except BaseException:
            # Only remove the new output this function just created.
            try:
                output.unlink()
            except OSError:
                print(f"WARNING: incomplete output could not be removed: {output}")
            raise
    raise SafetyStop("Too many existing outputs with the same base filename.")


def choose_destination(source: Path, source_cloud: bool) -> tuple[Path, bool]:
    kind = choose(
        "\nSave to [1] local folder or [2] OneDrive/SharePoint synced folder "
        f"(default {'2' if source_cloud else '1'}): ",
        {"1", "2"},
        "2" if source_cloud else "1",
    )
    cloud = kind == "2"
    if cloud:
        grant_synced_access("saving the result")
    default = source.parent if cloud == source_cloud else None
    prompt = "Output folder, or browse"
    if default:
        prompt += f" (Enter = {default})"
    folder = input_path(prompt + ": ", directory=True, default=default)
    return folder, cloud


def process_pdf(
    working: Path,
    temp: Path,
) -> tuple[Path, list[PageState], list[PageState], set[int]] | None:
    password = ""
    with fitz.open(str(working)) as doc:
        if pdf_encrypted(doc):
            password = getpass("PDF owner password (not displayed or saved): ")
    pdf_preflight(working, password)
    baseline = scan_pdf(working, password)
    eligible = {n for n, p in enumerate(baseline, 1) if p.blank}
    deleted = select_pages(baseline, eligible, working)
    if not deleted:
        return None
    print("\nDeleting approved pages in a temporary copy; rechecking EVERY kept page.")
    output = temp / "verified.pdf"
    after = delete_pdf(working, output, baseline, deleted, password)
    return output, baseline, after, deleted


def process_word(
    working: Path,
    temp: Path,
) -> tuple[Path, list[PageState], list[PageState], set[int]] | None:
    word_preflight(working)
    with WordSession() as word:
        word.open(working)
        preview = temp / "word_before.pdf"
        word.export_pdf(preview)
        baseline = scan_pdf(preview)
        if word.doc.ComputeStatistics(2) != len(baseline):
            raise SafetyStop("Word's page count differs from its PDF export. Stopping.")
        plans = {}
        for number, state in enumerate(baseline, 1):
            if not state.blank:
                continue
            try:
                plans[number] = word.plan(number, len(baseline))
            except SafetyStop as exc:
                print(f"Blank page {number} must be kept: {exc}")
        deleted = select_pages(baseline, set(plans), preview)
        if not deleted:
            return None
        print("\nDeleting approved ranges in the temporary Word copy.")
        word.apply([plans[n] for n in deleted])
        # Verify the saved, CLOSED and REOPENED Word document, not just in-memory state.
        word.close_document()
        word.open(working, read_only=True)
        check_pdf = temp / "word_after.pdf"
        word.export_pdf(check_pdf)
        print("Rechecking EVERY kept page of the reopened Word document.")
        after = verify_pdf(baseline, check_pdf, deleted)
        word.close_document()
    return working, baseline, after, deleted


def main() -> int:
    if any(x in {"-h", "--help"} for x in sys.argv[1:]):
        print(__doc__)
        return 0
    if sys.argv[1:]:
        print("Run without arguments for prompts, or use --help.")
        return 2
    print("REMOVE BLANK PAGES | PDF and Word | Original always preserved")
    print("Type cancel at any prompt to stop. Deletion defaults to SKIP.")
    require_dependencies()
    source_cloud = (
        choose(
            "\nFile location: [1] local device/drive or [2] OneDrive/SharePoint "
            "synced folder (default 1): ",
            {"1", "2"},
            "1",
        )
        == "2"
    )
    if source_cloud:
        grant_synced_access("reading the source")
    source = input_path("\nFull file path, or type browse: ")
    is_word = source.suffix.lower() in {".doc", ".docx"}
    require_dependencies(word=is_word)
    print(f"\nSource: {source}")
    print("Blank detection is strict: noisy scans and pages with any content are kept.")
    original_hash = file_hash(source)
    # No derived files are written to a synced folder until verification succeeds.
    with tempfile.TemporaryDirectory(
        prefix="blank_page_remover_", ignore_cleanup_errors=True
    ) as temporary:
        temp = Path(temporary)
        working = temp / ("working" + source.suffix.lower())
        shutil.copyfile(source, working)
        if file_hash(working) != original_hash:
            raise SafetyStop("Source changed while copying or the copy is incomplete.")
        result = process_word(working, temp) if is_word else process_pdf(working, temp)
        if result is None:
            print("\nSkipped. No result file created; original unchanged.")
            return 0
        verified, before, after, deleted = result
        print(f"\nVERIFIED: {len(before)} pages -> {len(after)} pages.")
        print(f"Removed {len(deleted)} approved blank page(s).")
        remaining = [n for n, p in enumerate(after, 1) if p.blank]
        print(f"Remaining blank pages: {len(remaining)}")
        if remaining:
            print(
                "Kept blank pages, using output page numbers: "
                + ", ".join(map(str, remaining))
            )
        while True:
            folder, cloud = choose_destination(source, source_cloud)
            print(f"Output folder: {folder}")
            if not yes("Save the verified copy here"):
                if yes("Choose another folder"):
                    continue
                raise Cancelled()
            try:
                saved = publish_verified_copy(verified, folder, source, original_hash)
                break
            except OSError as exc:
                print(f"Could not save to that folder: {exc}")
                if not yes("Try another folder while the verified copy is available"):
                    raise Cancelled() from exc
        print(f"\nSUCCESS: saved and read-back verified:\n{saved}")
        print("Original file was not overwritten.")
        if cloud:
            print(
                "Saved in the selected SYNCED FOLDER. Cloud upload is NOT verified.\n"
                "Check OneDrive sync status and the file in OneDrive/SharePoint online.\n"
                "Offline state, permissions, or organization policy can prevent upload."
            )
        print("Close any temporary PDF preview so temporary files can be cleaned up.")
    return 0


def entrypoint() -> int:
    try:
        return main()
    except (Cancelled, KeyboardInterrupt, EOFError):
        print("\nCancelled. The original was not overwritten.")
        return 0
    except SafetyStop as exc:
        print(f"\nSTOPPED SAFELY: {exc}")
        print("Original not overwritten. No successful result is being reported.")
        return 1
    except Exception as exc:
        print(f"\nCould not finish: {type(exc).__name__}: {exc}")
        print(
            "Original not overwritten. Check file access, Python dependencies, "
            "and (for Word) installed desktop Word; close password/protection dialogs."
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(entrypoint())
