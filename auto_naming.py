# auto_naming.py
"""Decides which Excel file a photo belongs to, from the folder it sits in.
No Streamlit and no SharePoint calls in here, so it can be tested on its own.

Photos are uploaded to:
    Whiteboard Photos / Project Name / [Tower #] / Elevation / Drop

  - The first folder (the project) is the output folder:  <Excel landing folder> / Project Name /
  - The Excel file is named from the whole path:  Project - Tower # - Elevation - Drop.xlsx
    (a project with a single tower has no Tower folder, so that part is simply left out)
  - A photo placed directly in a project folder goes into a file named after the project.
  - A photo placed loose in "Whiteboard Photos" goes into "Unsorted.xlsx" in the landing folder.

The number of folder levels is not fixed: everything below the project is joined with " - "."""
import re

UNSORTED_NAME = "Unsorted"


def clean_part(text):
    """A folder name with extra spaces removed: ' Tower   1 ' -> 'Tower 1'."""
    return re.sub(r"\s+", " ", str(text or "")).strip()


def target_for_folder(relative_parts, sanitize_filename, sanitize_foldername):
    """`relative_parts` are the folder names below "Whiteboard Photos" (a list, [] for a
    loose photo). The two sanitize functions turn a typed name into a safe file name
    (with .xlsx) / a safe folder name, or None when nothing usable is left.
    Returns (output_subfolder, excel_filename); the subfolder is "" for the landing folder."""
    parts = [clean_part(p) for p in relative_parts]
    parts = [p for p in parts if p]
    unsorted_file = sanitize_filename(UNSORTED_NAME)
    if not parts:
        return "", unsorted_file

    project = sanitize_foldername(parts[0])
    if not project:
        return "", unsorted_file

    stem = " - ".join([project] + parts[1:])
    return project, (sanitize_filename(stem) or unsorted_file)
