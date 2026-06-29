"""
file_mover.py
-------------
Handles COPYING files to their destination category folder.
Original files are KEPT in the source folder (copy, not move).
Creates folders if they don't exist.
Handles duplicate filenames by appending _1, _2, etc.
Fully offline — standard Python only.
"""

import os
import shutil
from pathlib import Path
from datetime import datetime


class FileMover:
    def __init__(self, base_folder: str):
        """
        Args:
            base_folder: Root folder where all category subfolders live.
                         e.g. "C:\\Users\\Paulraj\\Documents\\Railway Files"
        """
        self.base_folder = Path(base_folder)

    def _ensure_folder(self, folder: Path) -> Path:
        """Create folder (and parents) if it doesn't exist."""
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    def _resolve_destination(self, target_folder: Path, filename: str) -> Path:
        """
        Return a final destination path that doesn't clash with existing files.
        If 'report.pdf' exists, tries 'report_1.pdf', 'report_2.pdf', etc.
        """
        dest = target_folder / filename
        if not dest.exists():
            return dest

        stem = Path(filename).stem
        suffix = Path(filename).suffix
        counter = 1
        while True:
            new_name = f"{stem}_{counter}{suffix}"
            dest = target_folder / new_name
            if not dest.exists():
                return dest
            counter += 1

    def move(self, source_path: str, category: str) -> dict:
        """
        COPY a file to the correct category folder under base_folder.
        Original file is kept in place (copy, not move/cut).

        Args:
            source_path: Full path to the file to be copied.
            category: Category name (becomes the subfolder name).

        Returns:
            dict with:
                - success (bool)
                - destination (str): Final path where file was copied
                - was_renamed (bool): True if filename had to be changed to avoid clash
                - original_filename (str)
                - final_filename (str)
                - error (str): Error message if success is False
        """
        source = Path(source_path)

        if not source.exists():
            return {
                "success": False,
                "destination": None,
                "was_renamed": False,
                "original_filename": source.name,
                "final_filename": None,
                "error": f"Source file not found: {source_path}"
            }

    def find_matching_folder(self, category: str) -> tuple:
        """
        Scan base_folder for matches.
        Returns: (folder_name, matched_existing)
        """
        if not self.base_folder.exists():
            return self._sanitise_folder_name(category), False

        target_clean = self._sanitise_folder_name(category).lower().strip()
        if not target_clean:
            return "MISCELLANEOUS FILES", False

        try:
            existing_dirs = [d.name for d in self.base_folder.iterdir() if d.is_dir()]
        except Exception:
            existing_dirs = []

        # 1. Exact match (case-insensitive)
        for d in existing_dirs:
            if d.lower().strip() == target_clean:
                return d, True
                
        # 2. Substring match (starts with / contains)
        # Avoid matching extremely short folder names to prevent false positives
        for d in existing_dirs:
            d_clean = d.lower().strip()
            if len(target_clean) >= 3 or len(d_clean) >= 3:
                if target_clean in d_clean or d_clean in target_clean:
                    return d, True

        return self._sanitise_folder_name(category), False

    def move(self, source_path: str, category: str) -> dict:
        """
        COPY a file to the correct category folder under base_folder.
        Original file is kept in place (copy, not move/cut).

        Args:
            source_path: Full path to the file to be copied.
            category: Category name (becomes the subfolder name).

        Returns:
            dict with:
                - success (bool)
                - destination (str): Final path where file was copied
                - was_renamed (bool): True if filename had to be changed to avoid clash
                - original_filename (str)
                - final_filename (str)
                - matched_existing (bool)
                - resolved_category (str)
                - error (str): Error message if success is False
        """
        source = Path(source_path)

        if not source.exists():
            return {
                "success": False,
                "destination": None,
                "was_renamed": False,
                "original_filename": source.name,
                "final_filename": None,
                "matched_existing": False,
                "resolved_category": category,
                "error": f"Source file not found: {source_path}"
            }

        # Sanitise and find matching folder
        resolved_category, matched_existing = self.find_matching_folder(category)
        target_folder = self.base_folder / resolved_category
        self._ensure_folder(target_folder)

        destination = self._resolve_destination(target_folder, source.name)
        was_renamed = destination.name != source.name

        try:
            shutil.copy2(str(source), str(destination))  # COPY — keeps original
            return {
                "success": True,
                "destination": str(destination),
                "was_renamed": was_renamed,
                "original_filename": source.name,
                "final_filename": destination.name,
                "matched_existing": matched_existing,
                "resolved_category": resolved_category,
                "error": None
            }
        except PermissionError as e:
            return {
                "success": False,
                "destination": None,
                "was_renamed": False,
                "original_filename": source.name,
                "final_filename": None,
                "matched_existing": matched_existing,
                "resolved_category": resolved_category,
                "error": f"Permission denied: {e}"
            }
        except Exception as e:
            return {
                "success": False,
                "destination": None,
                "was_renamed": False,
                "original_filename": source.name,
                "final_filename": None,
                "matched_existing": matched_existing,
                "resolved_category": resolved_category,
                "error": str(e)
            }

    def _sanitise_folder_name(self, name: str) -> str:
        """Remove characters that are invalid in Windows folder names."""
        invalid = r'\/:*?"<>|'
        for ch in invalid:
            name = name.replace(ch, "_")
        return name.strip()

    def get_category_folder(self, category: str) -> str:
        """Return the full path to a category folder (may not exist yet). Uses smart matching."""
        resolved, _ = self.find_matching_folder(category)
        return str(self.base_folder / resolved)

    def list_categories(self) -> list:
        """Return list of existing category folder names under base_folder."""
        if not self.base_folder.exists():
            return []
        return [
            d.name for d in self.base_folder.iterdir()
            if d.is_dir()
        ]


# ---------- Quick test ----------
if __name__ == "__main__":
    import tempfile, os
    with tempfile.TemporaryDirectory() as tmpdir:
        base = os.path.join(tmpdir, "Railway Files")
        mover = FileMover(base)

        # Create a dummy file
        dummy = os.path.join(tmpdir, "TDR_2024_Test.pdf")
        with open(dummy, "w") as f:
            f.write("test content")

        result = mover.move(dummy, "Tenders")
        print("Move result:", result)
        print("File exists:", os.path.exists(result["destination"]))
