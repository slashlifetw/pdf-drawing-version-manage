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
    Singed version:     XX#-#-XXX##-X####-[version]-Signed.pdf
    Searchable version: XX#-#-XXX##-X####-[version]-Searchable.pdf

Where ''[version]'' is either a 1-2 digit integer or a single uppercase letter.

Example filenames::

    HT0-1-CIC01-D5026-1-Searchable.pdf
    HT0-1-CIC01-C5023_A-Signed.pdf
    HT0-1-CIC01-D5038-A.pdf

Author: Michael; JIUN-AN, TSAI; 蔡濬安
Version: 1.3.1
Last Updated: 2026/03/13
"""

import datetime
import os
import re
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox


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
# Core logic
# ---------------------------------------------------------------------------

def find_files(
    search_path: str,
    target_path: str,
    file_format: str,
    file_name_formats: list[str],
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
    """
    file_registry: list[FileInfo] = []

    for root, _, files in os.walk(search_path):
        for filename in files:

            # Pre-filter: skip non-PDF files immediately
            if not filename.endswith(file_format):
                continue

            # Pattern matching: try each regex in turn
            match = None
            for regex in file_name_formats:
                match = re.fullmatch(regex, filename)
                if match:
                    break

            # Skip filenames that don't conform to any known format
            if not match:
                continue

            # Extract structured matadata from the regex groups
            now_name = match.group(2)     # Base drawing number
            now_version = match.group(3)  # Version token (digit(s) or letter)
            now_groups = match.groups()

            # Group index 3 is only present for ''/ Singed/ Searchable variants
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

    # Copy all surviving files to the output folder
    for entry in file_registry:
        copy_file(entry.file_path, target_path, entry.full_filename)


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

def prepare_directories() -> tuple[str, str]:
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
    # Initiallise a hidden Tk root window (required by tkinter dialogs)
    window_root = tk.Tk()
    window_root.withdraw()

    # the user selects the root search folder
    choose_folder_path = filedialog.askdirectory(title="選擇要收尋的資料夾")

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

    1. Prompt the user to choose a root directory via :func:`prepare_directories`.
    2. Run :func:`find_files` to collect and copy the latest drawing versions.
    3. Display a completion message with the elapsed wall-clock time.

    Any un-handled exception is caught and surfeaced to the user through a
    ''tkinter'' error dialog.
    """
    try:
        # Obtain search root and freshly created output floder
        search_path, target_path = prepare_directories()
        file_format = 'pdf'

        # Perform the main extraction logic
        start_time = time.time()
        find_files(search_path, target_path, file_format, FILE_NAME_FORMATS)

        # Report completion time
        elapsed_time = round(time.time() - start_time, 2)
        print(f'執行完成，耗時：{elapsed_time} sec')
        messagebox.showinfo('Sucess', f'程式執行完成\n耗時:{elapsed_time} sec')

    except Exception as exc:
        messagebox.showerror('Error', f'程式執行發生錯誤:{exc}')


if __name__ == '__main__':
    main()
