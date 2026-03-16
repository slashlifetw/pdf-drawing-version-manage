"""Automated Latest Version Extraction Tool for Design Drawing PDFs.

This module automates the extraction of the latest versions of PDF design
drawings from a designated directory tree. It is designed for construction
project environments where vendors submit drawings with incremental version
suffixes in standardized filenames.

Key capabilities:
    1. Recursively searches for PDF files matching a specific naming pattern.
    2. Compares version identifiers and retains only the latest version per
       drawing number.
    3. Resolves ties between same-version files using a quality suffix
       priority ranking (Searchable > Basic > Signed).
    4. Copies all filtered latest-version files into a newly created output
       folder within the source directory.

Supported file naming formats:

    Basic format:       XX#-#-XXX##-X####-[version].pdf
    Signed version:     XX#-#-XXX##-X####-[version]-Signed.pdf
    Searchable version: XX#-#-XXX##-X####-[version]-Searchable.pdf

Where ''[version]'' is either a 1-2 digit integer or a single uppercase letter.

Example filenames::

    HT0-1-CIC01-D5026-1-Searchable.pdf
    HT0-1-CIC01-C5023_A-Signed.pdf
    HT0-1-CIC01-D5038-A.pdf

Author: Michael; JIUN-AN, TSAI; 蔡濬安
Version: 1.4
Last Updated: 2026/03/16
"""

import datetime
import os
import re
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ---------------------------------------------------------------------------
# Module-level constants
# ---------------------------------------------------------------------------

# Regular expression patterns for each supported filename variant.
# Each pattern captures three groups:
#   Group 1: Full filename       (e.g. "HT0-1-CIC01-D5026-1-Searchable.pdf")
#   Group 2: Base drawing number (e.g. "HT0-1-CIC01-D5026")
#   Group 3: Version token       (numeric string or single letter, e.g. "3" or "A")
#   Group 4: Suffix label        (Signed / Searchable variants only)
FILE_NAME_FORMATS = [
    r'(([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])\.pdf)',
    r'(([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])-(Signed)\.pdf)',
    r'(([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])-(Searchable)\.pdf)'
]

# Quality ranking for same-version files.
# Higher value = preferred when version numbers are equal.
SUFFIX_PRIORITY = {
    'Searchable': 3,  # Searchable version
    '': 2,            # Basic version
    'Signed': 1       # Signed version
}


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class FileInfo:
    """File information container for a single tracked PDF drawing file.

    When multiple files share the same drawings name, only one
    ''FileInfo'' is kept in the tracking list; its attributes
    are overwritten whenever a superior version is encountered.

    Attributes:
        base_name (str): Base name of the file (without version and suffix)
            e.g. ''"HT0-1-CIC01-D5038"''.
        version (str): Raw version token as extracted from the filename,
            e.g. ``"3"`` or ``"B"``.
        suffix (str): File suffix, one of ``"Searchable"``, ``"Signed"``,
            or ``""`` for files.
        file_path (str): Absolute filesystem path to the physical PDF file.
        full_filename (str): Complete filename,
            e.g. ''"HT0-1-CIC01-D5026-1-Searchable.pdf"''.
    """

    def __init__(
        self,
        base_name: str,
        version: str,
        suffix: str,
        file_path: str,
        full_filename: str,
    ) -> None:
        """Initialise a FileInfo instance with all required attributes.

        Args:
            base_name: Drawing identifier without version/suffix component.
            version: Version token extracted from the filename.
            suffix: File suffix label (``"Searchable"``, ``"Signed"``,
                or empty string).
            file_path: Absolute path to the PDF file on disk.
            full_filename: Complete filename string including the extension.
        """
        self.base_name = base_name
        self.version = version
        self.suffix = suffix
        self.file_path = file_path
        self.full_filename = full_filename

    def update_inf(
        self,
        new_version: str,
        new_suffix: str,
        new_file_path: str,
        new_full_filename: str,
    ) -> None:
        """Replace stored metadata with data from a superior file version.

        Called when the comparison logic determines that a newly discovered
        file should supersede the currently tracked entry.

        Args:
            new_version: Updated version token.
            new_suffix: Updated quality suffix.
            new_file_path: Absolute path to the superior file.
            new_full_filename: Complete filename of the superior file.
        """
        self.version = new_version
        self.suffix = new_suffix
        self.file_path = new_file_path
        self.full_filename = new_full_filename


# ---------------------------------------------------------------------------
# ★ Progress window — 全新新增
# ---------------------------------------------------------------------------

class ProgressWindow:
    """A live progress window displayed during the file extraction workflow.

    Shows real-time status while find_files() is running, including the
    current directory being scanned, the number of files processed, and
    a progress bar for the copy phase.

    The progress bar runs in indeterminate (marquee) mode during the scan
    phase and automatically switches to determinate mode when copying begins.

    The window disables manual closing; it must be dismissed programmatically
    by calling close() upon completion or when an exception is raised.

    Attributes:
        root (tk.Toplevel): The Toplevel window instance for this dialog.
        _stage_var (tk.StringVar): Text variable for the current stage label.
        _folder_var (tk.StringVar): Text variable for the current scan folder.
        _file_var (tk.StringVar): Text variable for the scanned/matched file
            counts.
        _bar (ttk.Progressbar): Progress bar widget; indeterminate during
            scanning, determinate during copying.
    """

    def __init__(self, parent: tk.Tk) -> None:
        """Initialise and display the progress window.

        Args:
            parent: The hidden root Tk window to attach this Toplevel to.
        """
        self.root = tk.Toplevel(parent)
        self.root.title('圖說最新版本抓取工具')
        self.root.resizable(False, False)
        self.root.protocol('WM_DELETE_WINDOW', lambda: None)

        self.root.geometry('420x180')
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 420) // 2
        y = (self.root.winfo_screenheight() - 180) // 2
        self.root.geometry(f'420x180+{x}+{y}')

        frame = tk.Frame(self.root, padx=24, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        self._stage_var = tk.StringVar(value='正在初始化...')
        tk.Label(frame, textvariable=self._stage_var,
                 font=('Microsoft JhengHei', 10, 'bold'),
                 anchor='w').pack(fill=tk.X)

        self._folder_var = tk.StringVar(value='')
        tk.Label(frame, textvariable=self._folder_var,
                 font=('Microsoft JhengHei', 9), fg='gray',
                 anchor='w', wraplength=370,
                 justify='left').pack(fill=tk.X, pady=(4, 0))

        self._file_var = tk.StringVar(value='')
        tk.Label(frame, textvariable=self._file_var,
                 font=('Microsoft JhengHei', 9), fg='gray',
                 anchor='w').pack(fill=tk.X, pady=(2, 10))

        self._bar = ttk.Progressbar(frame, mode='indeterminate', length=370)
        self._bar.pack(fill=tk.X)
        self._bar.start(12)

        self.root.update()

    def set_stage(self, text: str) -> None:
        """Update the stage label text.

        Args:
            text: Description of the current execution stage,
                e.g. 'Scanning drawing folders...'.
        """
        self._stage_var.set(text)
        self.root.update()

    def set_folder(self, folder: str) -> None:
        """Update the currently scanned folder path.

        Displays only the last two path components to keep the label concise.

        Args:
            folder: Absolute path of the folder currently being walked.
        """
        parts = folder.replace('\\', '/').split('/')
        short = '/'.join(parts[-2:]) if len(parts) >= 2 else folder
        self._folder_var.set(f'📁 {short}')
        self.root.update()

    def set_file_count(self, scanned: int, matched: int) -> None:
        """Update the scanned and matched file counters.

        Args:
            scanned: Total number of PDF files scanned so far.
            matched: Number of files that matched a known naming format.
        """
        self._file_var.set(f'已掃描 {scanned} 個 PDF，符合格式 {matched} 個')
        self.root.update()

    def switch_to_copy(self, total: int) -> None:
        """Switch the progress bar from indeterminate to determinate mode.

        Called once scanning is complete and copying is about to begin.

        Args:
            total: Total number of files to be copied; sets the progress
                bar maximum.
        """
        self._bar.stop()
        self._bar.config(mode='determinate', maximum=total, value=0)
        self.set_stage(f'正在複製最新版本（共 {total} 份）...')
        self.root.update()

    def increment_copy(self) -> None:
        """Advance the progress bar by one step after each file is copied."""
        self._bar.step(1)
        self.root.update()

    def close(self) -> None:
        """Destroy the progress window and release its resources."""
        self.root.destroy()


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def find_files(
    search_path: str,
    target_path: str,
    file_format: str,
    file_name_formats: list[str],
    progress: ProgressWindow | None = None,
) -> None:
    """Search *search_path* recursively and copy the latest drawing versions.

    For every PDF whose filename matches one of the patterns in
    *file_name_formats*, the function:

    1. Parses the base drawing number, version token, and quality suffix.
    2. Checks whether a file with the same base name has already been seen.
    3. If not, registers a new ``FileInfo`` entry in the local registry.
    4. If a prior entry exists, delegates to :func:`_compare_and_update` to
       decide whether the new file should replace the stored one.

    After the full walk completes, every surviving entry is copied to
    *target_path* via :func:`copy_file`.

    Version comparison rules (in descending priority):

    * Numeric versions always beat alphabetic versions (``"3"`` > ``"B"``).
    * Among two numeric versions, the larger integer wins.
    * Among two alphabetic versions, the lexicographically larger letter wins.
    * For equal versions, ``SUFFIX_PRIORITY`` determines the winner.

    Args:
        search_path: Root directory from which the recursive search begins.
        target_path: Destination directory where winning files are copied.
        file_format: File extension to filter on, without the dot (e.g. ``"pdf"``).
        file_name_formats: List of regex strings used to validate and parse
            candidate filenames.
        progress: Optional ProgressWindow instance for live status updates.
              Pass ``None`` to run silently (default).
    """
    file_registry: list[FileInfo] = []
    scanned = 0
    matched = 0

    for root, _, files in os.walk(search_path):

        if progress:
            progress.set_folder(root)

        for filename in files:

            # Pre-filter: skip non-PDF files immediately
            if not filename.endswith(file_format):
                continue

            scanned += 1

            # Pattern matching: try each regex in turn
            match = None
            for regex in file_name_formats:
                match = re.fullmatch(regex, filename)
                if match:
                    break

            # Skip filenames that don't conform to any known format
            if not match:
                continue

            matched += 1

            if progress:
                progress.set_file_count(scanned, matched)

            # Extract structured matadata from the regex groups
            now_name = match.group(2)     # Base drawing number
            now_version = match.group(3)  # Version token (digit(s) or letter)
            now_groups = match.groups()

            # Group index 3 is only present for ''/ Signed/ Searchable variants
            now_suffix = now_groups[3] if len(now_groups) == 4 and now_groups[3] else ''

            now_full_name = match.group(1)
            now_file_path = os.path.join(root, filename)

            # Search for an existing entry with the same drawing base name
            existing_entry = None
            for entry in file_registry:
                if entry.base_name == now_name:
                    existing_entry = entry
                    break

            if existing_entry is None:
                # No entry for this drawing yet - register it
                file_registry.append(
                    FileInfo(now_name,
                             now_version,
                             now_suffix,
                             now_file_path,
                             now_full_name)
                )
            else:
                # An entry already exists; determine whether to replace it
                _compare_and_update(
                    entry, now_version, now_suffix,
                    now_file_path, now_full_name
                )

    if progress:
        progress.switch_to_copy(len(file_registry))

    # Copy all surviving files to the output folder
    for entry in file_registry:
        copy_file(entry.file_path, target_path, entry.full_filename)
        if progress:
            progress.increment_copy()


def _compare_and_update(
    existing: FileInfo,
    now_version: str,
    now_suffix: str,
    now_file_path: str,
    now_full_filename: str,
) -> None:
    """Compare a newly found file against the currently registered entry.

    Normalises both version tokens and delegates to :func:`version_update`
    when the new file should replace the stored one.

    Comparison logic (evaluated in order):

    1. **Both numeric** – new wins if its integer value >= existing integer.
    2. **Existing alphabetic, new numeric** – numeric always supersedes.
    3. **Both alphabetic** – new wins if its letter >= existing letter.
    4. **Existing numeric, new alphabetic** – keep existing (no action).

    Args:
        existing: The ``FileInfo`` instance currently stored in the registry.
        now_version: Version token of the newly discovered file.
        now_suffix: Quality suffix of the newly discovered file.
        now_file_path: Absolute path to the newly discovered file.
        now_full_filename: Complete filename of the newly discovered file.
    """
    # Normalise version token for comparsion
    origin_version_is_numeric = existing.version.isdigit()
    now_version_is_numeric = now_version.isdigit()

    origin_version_comparable = int(existing.version) if origin_version_is_numeric else existing.version
    now_version_comparable = int(now_version) if now_version_is_numeric else now_version

    if isinstance(origin_version_comparable, int) and isinstance(now_version_comparable, int):
        # Both numeric: larger integer (or same version with better suffix) wins
        if now_version_comparable >= origin_version_comparable:
            version_update(existing, now_version_comparable, now_suffix, now_file_path, now_full_filename, origin_version_comparable)

    elif not origin_version_is_numeric and now_version_is_numeric:
        # Numeric always supersedes alphabetic
        version_update(existing, now_version, now_suffix, now_file_path, now_full_filename, existing.version)

    elif not origin_version_is_numeric and not now_version_is_numeric:
        if now_version_comparable >= origin_version_comparable:
            version_update(existing, now_version, now_suffix, now_file_path, now_full_filename, existing.version)
    # Remaining case (existing numeric, new alphabetic): keep existing – no action needed.


def version_update(
    existing_entry: FileInfo,
    new_version,
    new_suffix: str,
    new_file_path: str,
    new_full_filename: str,
    origin_version,
) -> None:
    """Overwrite a registered ''FileInfo'' entry if the new file is superior.

    Evaluates two sub-cases:

    * **Version greater** - replacement.
    * **Version equal** - replacement only when *new_suffix* has a higher
      ``SUFFIX_PRIORITY`` value than the currently stored suffix(origin_suffix).

    Args:
        existing_entry: The ``FileInfo`` object to overwrite.
        new_version: Comparable version token for the new file (may be
            ``int`` or ``str`` depending on caller).
        new_suffix: Quality suffix of the new file.
        new_file_path: Absolute path of the new file.
        new_full_filename: Complete filename of the new file.
        origin_version: Comparable version token for the existing file.
    """
    if new_version > origin_version:
        # New file has a higher version token
        existing_entry.update_inf(str(new_version), new_suffix, new_file_path, new_full_filename)
    elif not str(origin_version).isdigit() and str(new_version).isdigit():
        # Alphabetic -> numeric upgrade
        existing_entry.update_inf(str(new_version), new_suffix, new_file_path, new_full_filename)
    elif new_version == origin_version:
        # Same version; Compare the priority of suffix
        num_new_suffix = SUFFIX_PRIORITY.get(new_suffix)
        num_origin_suffix = SUFFIX_PRIORITY.get(existing_entry.suffix)
        if num_new_suffix > num_origin_suffix:
            existing_entry.update_inf(str(new_version), new_suffix, new_file_path, new_full_filename)


def copy_file(
    source_path: str,
    target_dir: str,
    filename: str,
) -> None:
    """Copy a single file form *source_path* into *target_dir*.

    The destination filename is always the same as the source filename

    Args:
        source_path: Absolute path of the file to copy.
        target_dir: Destination directory; must already exist.
        filename: Filename to use at the destination.
    """
    destination_path = os.path.join(target_dir, filename)
    print(f'  [COPY] {source_path}')
    print(f'      → {destination_path}')
    shutil.copyfile(source_path, destination_path)


# ---------------------------------------------------------------------------
# GUI helpers
# ---------------------------------------------------------------------------

def prepare_directories(parent: tk.Tk) -> tuple[str, str]:
    """Present a folder-selection dialog and prepare the output directory.

    Workflow:

    1. Opens a native folder-picker dialog for the user to choose the root
       search directory.
    2. Removes any pre-existing ``★★★最新版本(...)★★★`` subfolder to avoid
       stale results accumulating across runs.
    3. Creates a freshly named output folder stamped with today's date.

    Returns:
        A two-tuple ``(search_path, target_path)`` where *search_path* is the
        user-selected root directory and *target_path* is the newly created
        output subfolder.
    """

    # the user selects the root search folder
    choose_folder_path = filedialog.askdirectory(
        title="選擇要收尋的資料夾", parent=parent)
    if not choose_folder_path:
        raise RuntimeError('未選擇資料夾，程式結束。')

    today_date = datetime.date.today()

    # Remove stale output folders from previous runs
    for folder_name in os.listdir(choose_folder_path):
        full_folder_path = os.path.join(choose_folder_path, folder_name)
        if '最新版本' in folder_name and os.path.isdir(full_folder_path):
            shutil.rmtree(full_folder_path)
            messagebox.showinfo('Success', f'{folder_name}舊資料夾已刪除。')

    # Create the new dated output folder
    new_folder_name = f'★★★最新版本({today_date}更新)★★★'
    new_folder_path = os.path.join(choose_folder_path, new_folder_name)
    os.makedirs(new_folder_path)
    messagebox.showinfo('Success', f'已建立新資料夾:{new_folder_name}')

    return choose_folder_path, new_folder_path


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    """The full extraction workflow.

    Steps:

    1. Initialise a hidden Tk root window.
    2. Prompt the user to choose a root directory via prepare_directories.
    3. Display a ProgressWindow and run find_files to collect and copy
       the latest drawing versions.
    4. Display a completion message with the elapsed wall-clock time.

    Any un-handled exception is caught and surfeaced to the user through a
    ''tkinter'' error dialog.
    """
    root = tk.Tk()
    root.withdraw()

    progress = None

    try:
        # Obtain search root and freshly created output floder
        search_path, target_path = prepare_directories(root)
        progress = ProgressWindow(root)
        progress.set_stage('正在掃描圖說資料夾...')
        file_format = 'pdf'

        # Perform the main extraction logic
        start_time = time.time()
        find_files(search_path, target_path, file_format, FILE_NAME_FORMATS
                   , progress)

        progress.close()
        progress = None

        # Report completion time
        elapsed_time = round(time.time() - start_time, 2)
        print(f'執行完成，耗時：{elapsed_time} sec')
        messagebox.showinfo('Sucess', f'程式執行完成\n耗時:{elapsed_time} sec')

    except Exception as exc:
        if progress:
            progress.close()
        messagebox.showerror('Error', f'程式執行發生錯誤:{exc}')

    finally:
        root.destroy()


if __name__ == '__main__':
    main()
