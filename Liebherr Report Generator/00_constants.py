# 00_constants.py: global constants and generic helpers (paths, files). Must be executed FIRST in the Mechanical scripting console.

import os
import re
import shutil

# === Root paths, computed from PROJECT_DIR (already defined by AnsysReportGenerator_WPF.py before the execfile() of this file, = the "Report Generator" folder of the extension itself) ===
# No path is hardcoded for a specific machine/project: PROJECT_DIR is located via __file__ (see
# _prepare_environment in AnsysReportGenerator_WPF.py), i.e. the extension's own install folder -
# not the Ansys project currently open. DATA_ROOT and the folders below therefore live next to
# this script, shared by every Ansys project opened with this same extension, and are created
# automatically on first run if they don't already exist (see the ensure_folder_exists() calls
# at the bottom of this file).
DATA_ROOT = os.path.join(PROJECT_DIR, "data")

IMAGE_EXPORT_FOLDER = os.path.join(DATA_ROOT, "image_export")
CSV_EXPORT_FOLDER = os.path.join(DATA_ROOT, "csv_export")
EXPORT_3D_FOLDER = os.path.join(DATA_ROOT, "export_3D")

    # Unlike the folders above, legends are deliberately NOT shared across projects: engineers
    # save and edit their legend .xml files per Ansys project, in that project's own
    # "user_files/legend" folder (standard Ansys project folder) - this script only READS them,
    # never generates them. Located via ExtAPI.DataModel.Project.ProjectDirectory (the
    # "<Project>_files" folder of the project currently open in Mechanical): safe to call here
    # even though PROJECT_DIR itself no longer depends on ExtAPI, because 00_constants.py is only
    # ever loaded on the toolbar button click (HighFiveOut), by which point a Mechanical project
    # is necessarily open. Falls back to the shared data/legend folder (with a console warning)
    # if the project was never saved and ProjectDirectory is unavailable, so a missing/unsaved
    # project degrades gracefully instead of crashing the whole extension.
try:
    _ansys_project_directory = ExtAPI.DataModel.Project.ProjectDirectory
except Exception as _legend_dir_ex:
    _ansys_project_directory = None
    print "WARNING: unable to read ExtAPI.DataModel.Project.ProjectDirectory ({}): falling back to the shared legend folder.".format(str(_legend_dir_ex))

if _ansys_project_directory:
    LEGEND_FOLDER = os.path.join(_ansys_project_directory, "user_files", "legend")
else:
    print "WARNING: ExtAPI.DataModel.Project.ProjectDirectory is empty (save the Ansys project to use its own legends) - falling back to the shared legend folder."
    LEGEND_FOLDER = os.path.join(DATA_ROOT, "legend")

    # Folder for the working copy of the template (see PPTReportBuilder in 03_ppt_utils.py): the original template is never opened directly, to never risk being overwritten by an accidental Ctrl+S.
REPORT_OUTPUT_FOLDER = os.path.join(DATA_ROOT, "reports")

    # Directly in PROJECT_DIR (flat structure, no "templates" subfolder): unlike the folders above, this file cannot be created automatically if missing (see the warning further below).
TEMPLATE_PATH = os.path.join(PROJECT_DIR, "Master Template_def.pptx")

    # Company logo displayed in the sidebar credit card (see ReportGeneratorApp._load_logo,
    # SidebarLogoBitmap in the XAML): like TEMPLATE_PATH, cannot be created automatically if missing.
LOGO_PATH = os.path.join(PROJECT_DIR, "logo", "Liebherr-Emblem.png")


# === Custom layout indices of the PowerPoint template ===
LAYOUT_IMAGE_TABLE = 10    # title[2] / subtitle[4] / image[3] / table[1] / comment[8]
LAYOUT_TABLE_ONLY = 8      # title[1] / subtitle[3] / table[2]
LAYOUT_MESH_MULTI = 11     # images[5,6,7,8] (top) / tables[9,10,11,12] (bottom) -- indices on the generated SLIDE, not on the layout (see MESH_MULTI_*_SHAPE_INDICES below)

DEFAULT_IMAGE_WIDTH = 1920
DEFAULT_IMAGE_HEIGHT = 1920

# === Safeguard for displaying tables in PowerPoint ===
    # The CSV is always exported regardless of its size; only its insertion as a PowerPoint table is blocked beyond these limits (table becomes unreadable once inserted).
MAX_TABLE_ROWS = 50
MAX_TABLE_COLUMNS = 50

# === Mesh per isolated part (multi-image slide, see LAYOUT_MESH_MULTI) ===
    # Layout 11 of the template ("Custom Layout") contains, in SlideMaster.CustomLayouts,
    # an extra Table shape (not a placeholder) that is NOT inherited by slides created
    # from this layout: on layout.Shapes it occupies index 5 and shifts everything after it
    # (images at 6-9, tables at 10-13), but on the actually generated slide (report.presentation.Slides.AddSlide),
    # this shape is absent and everything shifts back by one (images at 5-8, tables at 9-12). The indices
    # below are the ones seen on the SLIDE (what the code actually manipulates), not on the layout.
MESH_MULTI_IMAGE_SHAPE_INDICES = [5, 6, 7, 8]
MESH_MULTI_TABLE_SHAPE_INDICES = [9, 10, 11, 12]
MAX_MESH_MULTI_BODIES = 4  # number of image/table slots available on this layout


def ensure_folder_exists(folder_path):
    """
    Does: creates the folder_path directory (and its parents) if it doesn't already exist.
    Depends on: os.path.exists / os.makedirs.
    Returns: nothing (side effect on the file system).
    """
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


def safe_file_name(name):
    """
    Does: replaces characters forbidden in a Windows file name (including "/" and "\\") with an underscore.
    Depends on: the re module (regex).
    Returns: str, the cleaned name, usable as-is in a file path.
    """
    # Without this cleanup, a Mechanical name like "Part/Solid" creates a fake subfolder ("Mesh_Part\Solid.csv") and the write fails.
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "object"


def get_unique_file_path(folder, base_name, extension):
    """
    Does: builds a file path that does not collide with an existing file, adding an incremental suffix if needed.
    Depends on: os.path.exists / os.path.join.
    Returns: str, an absolute path guaranteed not to exist at the time of the call.
    """
    path = os.path.join(folder, base_name + extension)
    counter = 1
    while os.path.exists(path):
        path = os.path.join(folder, base_name + "_" + str(counter) + extension)
        counter += 1
    return path


def list_data_cleanup_folders():
    """
    Does: lists the direct subfolders of DATA_ROOT that can be offered for cleanup (Files tab),
    excluding the legend folder (never affected by cleanup - unlike the rest,
    these are not exports but configuration files reused from one generation to the next).
    Depends on: DATA_ROOT, LEGEND_FOLDER, os.listdir/os.path.isdir.
    Returns: list of tuples (display_name, absolute_path), sorted by name (empty if DATA_ROOT doesn't exist).
    """
    if not os.path.isdir(DATA_ROOT):
        return []
    legend_name = os.path.basename(os.path.normpath(LEGEND_FOLDER))
    folders = []
    for name in os.listdir(DATA_ROOT):
        path = os.path.join(DATA_ROOT, name)
        if os.path.isdir(path) and name != legend_name:
            folders.append((name, path))
    return sorted(folders, key=lambda item: item[0].lower())


def get_folder_stats(folder_path):
    """
    Does: computes the total size and file count of a folder (recursive, including subfolders).
    Depends on: os.walk, os.path.getsize.
    Returns: tuple (total_size_bytes, file_count) - (0, 0) if the folder doesn't exist.
    """
    total_size = 0
    file_count = 0
    if not os.path.isdir(folder_path):
        return total_size, file_count
    for root, _dirs, files in os.walk(folder_path):
        for name in files:
            try:
                total_size += os.path.getsize(os.path.join(root, name))
                file_count += 1
            except OSError:
                pass
    return total_size, file_count


def format_folder_size(size_bytes):
    """
    Does: formats a size in bytes into a human-readable string (bytes/KB/MB/GB).
    Depends on: nothing (pure calculation).
    Returns: str, the formatted size (e.g.: "12.4 MB").
    """
    size = float(size_bytes)
    for unit in ("bytes", "KB", "MB"):
        if size < 1024.0:
            if unit == "bytes":
                return "{} {}".format(int(size), unit)
            return "{:.1f} {}".format(size, unit)
        size /= 1024.0
    return "{:.1f} GB".format(size)


def clear_folder_contents(folder_path):
    """
    Does: deletes all the contents (files and subfolders) of a folder, without deleting the folder itself.
    Depends on: os.listdir, os.remove, shutil.rmtree.
    Returns: nothing (side effect on the file system; does nothing if the folder doesn't exist).
    """
    if not os.path.isdir(folder_path):
        return
    for name in os.listdir(folder_path):
        path = os.path.join(folder_path, name)
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
        except Exception as e:
            print "Unable to delete: {} ({})".format(path, str(e))


def clean_cell_text(text):
    """
    Does: normalizes the text of a Tabular Data pane cell for CSV export.
    Depends on: nothing (pure string processing).
    Returns: str, the cleaned text ("" if the input was None).
    """
    if text is None:
        return ""
    return text.replace("=", "").strip().rstrip(",").strip()


def to_csv_cell(value):
    """
    Does: converts any value (.NET unicode text, number, None) into a UTF-8-encoded str for csv.writer.
    Depends on: IronPython 2.7's unicode type.
    Returns: UTF-8-encoded str ("" if value is None).
    """
    # Some units returned by Mechanical (mm3, degree, micro...) contain special characters that crash the write if the encoding isn't explicitly fixed.
    if value is None:
        return ""
    if isinstance(value, unicode):
        return value.encode("utf-8")
    return str(value)


# First run on a new install: these storage folders don't exist yet,
# they are created here once and for all before the rest of the application uses them.
# LEGEND_FOLDER is NOT part of this: it is maintained manually by the engineer in the current
# Ansys project's "user_files" (see its definition above) - creating it automatically here would
# mask a genuine absence of legends instead of warning the user.
ensure_folder_exists(IMAGE_EXPORT_FOLDER)
ensure_folder_exists(CSV_EXPORT_FOLDER)
ensure_folder_exists(REPORT_OUTPUT_FOLDER)
ensure_folder_exists(EXPORT_3D_FOLDER)

# The template cannot be created automatically (a content file, not just a folder):
# we only warn in the console so the user immediately knows why generation
# would fail, without blocking the loading of the following modules.
if not os.path.isfile(TEMPLATE_PATH):
    print "WARNING: PowerPoint template not found at the expected location: " + TEMPLATE_PATH

# Same logic for the legend folder (see LEGEND_FOLDER above): not created automatically, we
# simply warn if the expected location in the current project's "user_files" doesn't exist.
if not os.path.isdir(LEGEND_FOLDER):
    print "WARNING: legend folder not found at the expected location: " + LEGEND_FOLDER

# Same logic for the logo (see LOGO_PATH above): non-blocking absence, just a
# warning (the imgLogo location in the XAML then simply stays empty).
if not os.path.isfile(LOGO_PATH):
    print "WARNING: logo not found at the expected location: " + LOGO_PATH
