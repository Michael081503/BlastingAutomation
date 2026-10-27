"""Save prepared files in a new local folder for each export."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from pathlib import Path

from automation import AutomationError
from exports import export_filename


PROJECT_DIRECTORY = Path(__file__).resolve().parent
MANILA_TIMEZONE = timezone(timedelta(hours=8), name="Asia/Manila")


def save_files_to_folder(
    files: list[tuple[str, bytes]],
    output_folder: str | Path,
    *,
    batch_name: str,
) -> Path:
    """Create a unique batch folder and save the CSV files without a ZIP.

    Relative paths are resolved from the project folder. Existing export folders
    and files are kept; same-second exports receive a numbered suffix.
    """
    if not files:
        raise AutomationError("There are no files to save. Prepare the exports first.")
    if not str(output_folder).strip():
        raise AutomationError("Enter an output folder before saving.")
    names = []
    for name, _ in files:
        extension = Path(name).suffix.lower()
        if extension != ".csv" or export_filename(name, extension) != name:
            raise AutomationError(f'Invalid output filename "{name}". Check export_config.py.')
        names.append(name.casefold())
    if len(set(names)) != len(names):
        raise AutomationError("Output filenames must be unique. Check export_config.py.")

    folder = None
    try:
        parent = Path(str(output_folder).strip()).expanduser()
        if not parent.is_absolute():
            parent = PROJECT_DIRECTORY / parent
        parent = parent.resolve()
        parent.mkdir(parents=True, exist_ok=True)
        label = re.sub(r"[^\w .-]+", "_", batch_name).strip(" .")[:80].rstrip(" .") or "export"
        timestamp = datetime.now(MANILA_TIMEZONE).strftime("%Y-%m-%d_%H-%M-%S")
        base_name = f"{label}_{timestamp}"
        number = 1
        while True:
            candidate = parent / (base_name if number == 1 else f"{base_name}_{number}")
            try:
                candidate.mkdir()
                folder = candidate
                break
            except FileExistsError:
                number += 1
        for name, data in files:
            with (folder / name).open("xb") as output:
                output.write(data)
    except (OSError, ValueError) as exc:
        detail = f" Some files may have been saved in {folder}." if folder else ""
        raise AutomationError(
            "Could not save the files. Check that the folder exists or can be created "
            "and that you have permission to write there." + detail
        ) from exc
    return folder
