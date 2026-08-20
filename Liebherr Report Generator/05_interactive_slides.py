# 05_interactive_slides.py: support logic for interactive selection (AnsysReportGenerator_WPF.py) - per-row configuration classes, collection of views/sections/analyses, and slide builders limited to the user's selection (instead of always processing a whole category like 04_slides.py). Depends on 00_constants.py, 01_data_export.py, 02_image_export.py, 03_ppt_utils.py, 04_slides.py (must be executed before this file).

import csv


def remove_stale_figures():
    """
    Does: removes leftover Figure objects from a previous report generation.
    Depends on: DataModel.GetObjectsByType/Remove, Transaction (Ansys.ACT.Mechanical).
    Returns: nothing (side effect: cleans up the tree before a new generation).
    """
    try:
        # Transaction(True) defers the tree/viewport refresh until the bulk deletion is complete.
        with Transaction(True):
            DataModel.Remove(DataModel.GetObjectsByType(DataModelObjectCategory.Figure))
    except Exception as e:
        print "Unable to delete existing figures: " + str(e)


# Column 2 (ViewOrientationType) corrected after a visual check in Mechanical:
# the part's orientation means the actual result of each enum does not match its
# natural name in the .NET API - e.g. ViewOrientationType.Back is what actually produces the X+ view.
BASIC_VIEW_ORIENTATIONS = [
    ("X+", "Back"),
    ("X-", "Front"),
    ("Y+", "Left"),
    ("Y-", "Right"),
    ("Z+", "Top"),
    ("Z-", "Bottom"),
    ("ISO", "Iso"),
]


def create_basic_views():
    """
    Does: creates 7 basic views (X+, X-, Y+, Y-, Z+, Z-, ISO) in the View Manager.
    Depends on: Ansys.Mechanical.DataModel.Enums.ViewOrientationType, ExtAPI.Graphics.Camera/ModelViewManager.
    Returns: list, the names of the views actually created.
    """
    from Ansys.Mechanical.DataModel.Enums import ViewOrientationType

    cam = ExtAPI.Graphics.Camera
    mvm = ExtAPI.Graphics.ModelViewManager

    # Transaction(True) defers the UI refresh until all 7 views are created: CreateView() only depends on the camera state, not on an already-displayed render.
    created = []
    with Transaction(True):
        for name, orientation_attr in BASIC_VIEW_ORIENTATIONS:
            try:
                orientation = getattr(ViewOrientationType, orientation_attr)
                cam.SetSpecificViewOrientation(orientation)
                cam.SetFit()
                mvm.CreateView(name)
                created.append(name)
                print "View created: " + name
            except Exception as e:
                print "Unable to create view {}: {}".format(name, str(e))

    return created


def export_object_3d_view(obj, directory):
    """
    Does: activates an object and exports its interactive 3D view (.avz) via the View Manager.
    Depends on: obj.Activate, ExtAPI.Graphics.ModelViewManager.Capture3DImage, get_unique_file_path/safe_file_name (00_constants.py).
    Returns: str, the path of the generated .avz file, or None on error.
    """
    try:
        obj.Activate()
        avz_path = get_unique_file_path(directory, safe_file_name(obj.Name), ".avz")
        ExtAPI.Graphics.ModelViewManager.Capture3DImage(avz_path)
        print "3D view exported: " + avz_path
        return avz_path
    except Exception as e:
        print "Unable to export 3D for {}: {}".format(obj.Name, str(e))
        return None


def collect_3d_exportable_objects(analysis):
    """
    Does: lists the objects to export in 3D for an analysis (simple results + Contact Tool/Bolt Tool from the Solution branch).
    Depends on: collect_all_results, collect_contact_tool_results, collect_bolt_tool_results.
    Returns: list, the Mechanical objects exportable to .avz for this analysis.
    """
    # Contact Tool / Bolt Tool from the Connections branch (definition, without a proper 3D result)
    # are intentionally excluded: only those from the Solution branch make sense here.
    objects = list(collect_all_results(analysis))
    objects.extend(collect_contact_tool_results(analysis))
    objects.extend(collect_bolt_tool_results(analysis))
    return objects


def export_all_3d_views(directory):
    """
    Does: exports to .avz the 3D view of all simple results and Contact/Bolt Tool (Solution branch) of all analyses in the project.
    Depends on: ensure_folder_exists (00_constants.py), collect_analyses, collect_3d_exportable_objects, export_object_3d_view.
    Returns: int, the number of .avz files actually exported.
    """
    ensure_folder_exists(directory)
    exported_count = 0
    for analysis in collect_analyses():
        for obj in collect_3d_exportable_objects(analysis):
            if export_object_3d_view(obj, directory):
                exported_count += 1
    return exported_count


NO_VIEW_LABEL = "-- Current view --"
NO_SECTION_LABEL = "-- No section --"


class SlideRowConfig(object):
    """
    Display configuration for ONE row of a selection list (a BC, a result, ...): object, view/section to apply before capture, optional steps.
    """

    def __init__(self, obj, analysis=None):
        """
        Does: initializes a row's configuration (default view/section/steps/scale/legend).
        Depends on: nothing (simple assignments).
        Returns: nothing (constructor).
        """
        self.obj = obj
        # None if the category is outside multi-analysis (BC, Contacts...) or the project is single-analysis: no suffix shown in that case (see analysis_suffix).
        self.analysis = analysis
        self.view_name = None
        self.section_name = None
        self.selected_steps = None       # None or empty list = no per-step handling
        self.step_display_mode = "individual"  # "individual" or "combined"
        self.scale_factor = 1.0          # deformation scale factor ("manual" mode only, 1.0 = no scaling)
        self.deformation_scale_mode = DEFAULT_DEFORMATION_SCALE_MODE  # "manual"/"auto_x1"/"auto_x2", see DEFORMATION_SCALE_MODE_OPTIONS
        self.legend_name = None          # legend name (see collect_legend_files), None = current/automatic legend
        self.contour_view = DEFAULT_CONTOUR_VIEW            # result color display mode (Isolines/SmoothContours/SolidFill/ContourBands)
        self.legend_orientation = DEFAULT_LEGEND_ORIENTATION  # legend orientation (Vertical/Horizontal)
        self.scoping_display = DEFAULT_SCOPING_DISPLAY       # scoping display (ScopedBodies/ResultOnly/AllBodies)
        self.configured = False          # becomes True once the "..." button has been confirmed (OK)


def analysis_suffix(row_config):
    """
    Does: builds the " (Analysis Name)" suffix shown to differentiate the same result between two analyses.
    Depends on: row_config.analysis (see collect_analyses and the *_multi collectors in this file).
    Returns: str, the formatted suffix, or an empty string if row_config.analysis is None.
    """
    if row_config.analysis is not None:
        return " ({})".format(row_config.analysis.Name)
    return ""


def build_row_display_name(row_config):
    """
    Does: builds the text shown for a selection row in the list.
    Depends on: row_config (obj, view_name, section_name, selected_steps, deformation_scale_mode, scale_factor, legend_name, contour_view, legend_orientation, scoping_display), analysis_suffix.
    Returns: str, the object's name followed by the chosen settings separated by " | ".
    """
    parts = [row_config.obj.Name + analysis_suffix(row_config)]
    if row_config.view_name:
        parts.append("view=" + row_config.view_name)
    if row_config.section_name:
        parts.append("section=" + row_config.section_name)
    if row_config.selected_steps:
        mode_label = "combined" if row_config.step_display_mode == "combined" else "individual"
        steps_label = ",".join(str(step) for step in row_config.selected_steps)
        parts.append("steps={} ({})".format(steps_label, mode_label))
    if row_config.deformation_scale_mode == "auto_x1":
        parts.append("scale=Auto x1")
    elif row_config.deformation_scale_mode == "auto_x2":
        parts.append("scale=Auto x2")
    elif row_config.scale_factor and row_config.scale_factor != 1.0:
        parts.append("scale=x{}".format(row_config.scale_factor))
    if row_config.legend_name:
        parts.append("legend=" + row_config.legend_name)
    if row_config.contour_view and row_config.contour_view != DEFAULT_CONTOUR_VIEW:
        parts.append("display=" + contour_view_label(row_config.contour_view))
    if row_config.legend_orientation and row_config.legend_orientation != DEFAULT_LEGEND_ORIENTATION:
        parts.append("legend_orientation=" + legend_orientation_label(row_config.legend_orientation))
    if row_config.scoping_display and row_config.scoping_display != DEFAULT_SCOPING_DISPLAY:
        parts.append("scoping=" + scoping_display_label(row_config.scoping_display))
    return " | ".join(parts)


def collect_views():
    """
    Does: lists the views saved in Mechanical's View Manager.
    Depends on: ExtAPI.Graphics.ModelViewManager.ExportModelViews, a temporary XML export, xml.etree.ElementTree.
    Returns: dict {name (str): index (int)}, empty if no views or on error.
    """
    views = {}
    try:
        view_manager = ExtAPI.Graphics.ModelViewManager
        xml_path = os.path.join(CSV_EXPORT_FOLDER, "_model_views_tmp.xml")
        # The View Manager cannot be inspected directly via the scripting API: a temporary XML export is used instead.
        view_manager.ExportModelViews(xml_path)
        tree = ET.parse(xml_path)
        for index, node in enumerate(list(tree.getroot())):
            if node.tag == "ModelView":
                views[node.attrib["Name"]] = index
    except Exception as e:
        print "View Manager views unavailable: " + str(e)
    return views


def collect_section_planes():
    """
    Does: lists the section planes already defined in the model.
    Depends on: ExtAPI.Graphics.SectionPlanes.
    Returns: list, the Section Plane objects found (empty on error).
    """
    planes = []
    try:
        section_planes = ExtAPI.Graphics.SectionPlanes
        for i in range(section_planes.Count):
            planes.append(section_planes[i])
    except Exception as e:
        print "Section planes unavailable: " + str(e)
    return planes


def section_plane_label(section_plane, index):
    """
    Does: builds a readable label for a section plane.
    Depends on: section_plane.Name.
    Returns: str, the section plane's name, or a generated name ("Section Plane N") if it has none.
    """
    try:
        if section_plane.Name:
            return section_plane.Name
    except Exception:
        pass
    return "Section Plane {}".format(index + 1)


def apply_view_if_exists(view_name, views):
    """
    Does: applies a View Manager view by name, if it still exists.
    Depends on: ExtAPI.Graphics.ModelViewManager.ApplyModelView, the views dict (see collect_views).
    Returns: nothing (side effect: changes the viewport view, or does nothing if absent).
    """
    if not view_name or view_name not in views:
        return
    try:
        ExtAPI.Graphics.ModelViewManager.ApplyModelView(views[view_name])
    except Exception as e:
        print "Unable to apply view '{}': {}".format(view_name, str(e))


def apply_section_plane(section_planes, section_labels, section_name):
    """
    Does: activates only the section plane designated by section_name, deactivates the others.
    Depends on: disable_all_section_planes, the index correspondence between section_planes and section_labels.
    Returns: nothing (side effect: changes the Active state of the section planes).
    """
    if not section_name:
        disable_all_section_planes(section_planes)
        return
    for i in range(len(section_planes)):
        try:
            section_planes[i].Active = (section_labels[i] == section_name)
        except Exception:
            pass


def disable_all_section_planes(section_planes):
    """
    Does: deactivates all the given section planes.
    Depends on: nothing (iterates over the given list).
    Returns: nothing (side effect: resets the section planes to a neutral state before/after capture).
    """
    for section_plane in section_planes:
        try:
            section_plane.Active = False
        except Exception:
            pass


DEFORMATION_SCALE_MODE_OPTIONS = [
    ("Manual (value below)", "manual"),
    ("Auto Scale x1", "auto_x1"),
    ("Auto Scale x2", "auto_x2"),
]
DEFAULT_DEFORMATION_SCALE_MODE = "manual"


def deformation_scale_mode_label(value):
    """
    Does: finds the label shown for a deformation_scale_mode value.
    Depends on: DEFORMATION_SCALE_MODE_OPTIONS.
    Returns: str, the matching label (the one for DEFAULT_DEFORMATION_SCALE_MODE if value is unknown).
    """
    for label, option_value in DEFORMATION_SCALE_MODE_OPTIONS:
        if option_value == value:
            return label
    return deformation_scale_mode_label(DEFAULT_DEFORMATION_SCALE_MODE)


def deformation_scale_mode_from_label(label):
    """
    Does: finds the deformation_scale_mode value associated with a label from DEFORMATION_SCALE_MODE_OPTIONS.
    Depends on: DEFORMATION_SCALE_MODE_OPTIONS.
    Returns: str, the matching value (DEFAULT_DEFORMATION_SCALE_MODE if the label is unknown).
    """
    for option_label, value in DEFORMATION_SCALE_MODE_OPTIONS:
        if option_label == label:
            return value
    return DEFAULT_DEFORMATION_SCALE_MODE


def apply_scale_factor(deformation_scale_mode, scale_factor):
    """
    Does: forces the deformation scale before image capture - either a manual factor
    (DeformationScaleMultiplier alone, original behavior), or one of Mechanical's two native
    "Auto Scale" presets (DeformationScaling forced to Auto + fixed multiplier 1 or 2).
    Depends on: ExtAPI.Graphics.ViewOptions.ResultPreference.DeformationScaling/DeformationScaleMultiplier,
        MechanicalEnums.Graphics.DeformationScaling (Ansys API, ambient enum), ExtAPI.Graphics.Redraw.
    Returns: nothing (side effect: changes the displayed scale, or does nothing in manual mode with scale_factor at 1.0).
    """
    try:
        vo = ExtAPI.Graphics.ViewOptions
        if deformation_scale_mode == "auto_x1":
            vo.ResultPreference.DeformationScaling = MechanicalEnums.Graphics.DeformationScaling.Auto
            vo.ResultPreference.DeformationScaleMultiplier = 1
            ExtAPI.Graphics.Redraw()
        elif deformation_scale_mode == "auto_x2":
            vo.ResultPreference.DeformationScaling = MechanicalEnums.Graphics.DeformationScaling.Auto
            vo.ResultPreference.DeformationScaleMultiplier = 2
            ExtAPI.Graphics.Redraw()
        elif scale_factor and scale_factor != 1.0:
            vo.ResultPreference.DeformationScaleMultiplier = float(scale_factor)
            ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Unable to apply scale factor: " + str(e)


def reset_scale_factor():
    """
    Does: resets the deformation scale to a neutral state (Manual mode, multiplier 1) after
    a capture with a custom value - including after an "Auto Scale" preset, so as not to
    leave MechanicalEnums.Graphics.DeformationScaling on Auto for the next capture.
    Depends on: ExtAPI.Graphics.ViewOptions.ResultPreference.DeformationScaling/DeformationScaleMultiplier,
        MechanicalEnums.Graphics.DeformationScaling (Ansys API, ambient enum), ExtAPI.Graphics.Redraw.
    Returns: nothing (side effect: resets the displayed deformation scale).
    """
    try:
        vo = ExtAPI.Graphics.ViewOptions
        vo.ResultPreference.DeformationScaling = MechanicalEnums.Graphics.DeformationScaling.UserDefined
        vo.ResultPreference.DeformationScaleMultiplier = 1
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Unable to reset scale factor: " + str(e)


CONTOUR_VIEW_OPTIONS = [
    ("ContourBands (default)", "ContourBands"),
    ("Isolines", "Isolines"),
    ("SmoothContours", "SmoothContours"),
    ("SolidFill", "SolidFill"),
]
DEFAULT_CONTOUR_VIEW = "ContourBands"


def contour_view_label(value):
    """
    Does: finds the label shown for a contour_view value.
    Depends on: CONTOUR_VIEW_OPTIONS.
    Returns: str, the matching label (the one for DEFAULT_CONTOUR_VIEW if value is unknown).
    """
    for label, option_value in CONTOUR_VIEW_OPTIONS:
        if option_value == value:
            return label
    return contour_view_label(DEFAULT_CONTOUR_VIEW)


def contour_view_from_label(label):
    """
    Does: finds the contour_view value associated with a label from CONTOUR_VIEW_OPTIONS.
    Depends on: CONTOUR_VIEW_OPTIONS.
    Returns: str, the matching value (DEFAULT_CONTOUR_VIEW if the label is unknown).
    """
    for option_label, value in CONTOUR_VIEW_OPTIONS:
        if option_label == label:
            return value
    return DEFAULT_CONTOUR_VIEW


def apply_contour_view(contour_view):
    """
    Does: applies the result color display mode (Isolines/Smooth Contours/Solid Fill/Contour Bands).
    Depends on: ExtAPI.Graphics.ViewOptions.ResultPreference.ContourView, ExtAPI.Graphics.Redraw (Ansys API).
    Returns: nothing (side effect on the viewport, or does nothing if contour_view is empty).
    """
    if not contour_view:
        return
    try:
        vo = ExtAPI.Graphics.ViewOptions
        # The members (Isolines/SmoothContours/SolidFill/ContourBands) are read from the TYPE of
        # the current value rather than imported explicitly: this is an ambient .NET enum, already
        # used this way elsewhere in the project (e.g. ModelColoring.ByMaterial).
        vo.ResultPreference.ContourView = getattr(vo.ResultPreference.ContourView, contour_view)
        # Redraw() is essential: changing this property via script does not refresh the
        # viewport on its own (same observation as for the legend, see reset_legend) - without this call, the image
        # exported right after stays on the old display mode.
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Unable to apply display mode '{}': {}".format(contour_view, str(e))


def reset_contour_view():
    """
    Does: resets the result color display mode to Contour Bands (neutral state) after a custom capture.
    Depends on: apply_contour_view, DEFAULT_CONTOUR_VIEW.
    Returns: nothing (side effect on the viewport).
    """
    apply_contour_view(DEFAULT_CONTOUR_VIEW)


LEGEND_ORIENTATION_OPTIONS = [
    ("Vertical (default)", "Vertical"),
    ("Horizontal", "Horizontal"),
]
DEFAULT_LEGEND_ORIENTATION = "Vertical"


def legend_orientation_label(value):
    """
    Does: finds the label shown for a legend_orientation value.
    Depends on: LEGEND_ORIENTATION_OPTIONS.
    Returns: str, the matching label (the one for DEFAULT_LEGEND_ORIENTATION if value is unknown).
    """
    for label, option_value in LEGEND_ORIENTATION_OPTIONS:
        if option_value == value:
            return label
    return legend_orientation_label(DEFAULT_LEGEND_ORIENTATION)


def legend_orientation_from_label(label):
    """
    Does: finds the legend_orientation value associated with a label from LEGEND_ORIENTATION_OPTIONS.
    Depends on: LEGEND_ORIENTATION_OPTIONS.
    Returns: str, the matching value (DEFAULT_LEGEND_ORIENTATION if the label is unknown).
    """
    for option_label, value in LEGEND_ORIENTATION_OPTIONS:
        if option_label == label:
            return value
    return DEFAULT_LEGEND_ORIENTATION


def apply_legend_orientation(legend_orientation):
    """
    Does: applies the viewport legend orientation (vertical/horizontal).
    Depends on: ExtAPI.Graphics.GlobalLegendSettings.LegendOrientation, LegendOrientationType (Ansys API, ambient enum), ExtAPI.Graphics.Redraw.
    Returns: nothing (side effect on the viewport, or does nothing if legend_orientation is empty).
    """
    if not legend_orientation:
        return
    try:
        ExtAPI.Graphics.GlobalLegendSettings.LegendOrientation = getattr(LegendOrientationType, legend_orientation)
        # Redraw() is essential: changing this property via script does not refresh the
        # viewport on its own (same observation as for the legend, see reset_legend) - without this call, the image
        # exported right after stays on the old orientation.
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Unable to apply legend orientation '{}': {}".format(legend_orientation, str(e))


def reset_legend_orientation():
    """
    Does: resets the legend orientation to Vertical (neutral state) after a custom capture.
    Depends on: apply_legend_orientation, DEFAULT_LEGEND_ORIENTATION.
    Returns: nothing (side effect on the viewport).
    """
    apply_legend_orientation(DEFAULT_LEGEND_ORIENTATION)


SCOPING_DISPLAY_OPTIONS = [
    ("ScopedBodies (default)", "ScopedBodies"),
    ("ResultOnly", "ResultOnly"),
    ("AllBodies", "AllBodies"),
]
DEFAULT_SCOPING_DISPLAY = "ScopedBodies"


def scoping_display_label(value):
    """
    Does: finds the label shown for a scoping_display value.
    Depends on: SCOPING_DISPLAY_OPTIONS.
    Returns: str, the matching label (the one for DEFAULT_SCOPING_DISPLAY if value is unknown).
    """
    for label, option_value in SCOPING_DISPLAY_OPTIONS:
        if option_value == value:
            return label
    return scoping_display_label(DEFAULT_SCOPING_DISPLAY)


def scoping_display_from_label(label):
    """
    Does: finds the scoping_display value associated with a label from SCOPING_DISPLAY_OPTIONS.
    Depends on: SCOPING_DISPLAY_OPTIONS.
    Returns: str, the matching value (DEFAULT_SCOPING_DISPLAY if the label is unknown).
    """
    for option_label, value in SCOPING_DISPLAY_OPTIONS:
        if option_label == label:
            return value
    return DEFAULT_SCOPING_DISPLAY


def apply_scoping_display(scoping_display):
    """
    Does: applies the scoping display mode (scoped bodies / result only / all bodies) before capture.
    Depends on: ExtAPI.Graphics.ViewOptions.ResultPreference.ScopingDisplay, MechanicalEnums.Graphics.ScopingDisplay (Ansys API, ambient enum), ExtAPI.Graphics.Redraw.
    Returns: nothing (side effect on the viewport, or does nothing if scoping_display is empty).
    """
    if not scoping_display:
        return
    try:
        vo = ExtAPI.Graphics.ViewOptions
        vo.ResultPreference.ScopingDisplay = getattr(MechanicalEnums.Graphics.ScopingDisplay, scoping_display)
        # Redraw() is essential: changing this property via script does not refresh the
        # viewport on its own (same observation as for ContourView/LegendOrientation) - without this call, the image
        # exported right after stays on the old display mode.
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Unable to apply scoping mode '{}': {}".format(scoping_display, str(e))


def reset_scoping_display():
    """
    Does: resets the scoping display mode to Scoped Bodies (neutral state) after a custom capture.
    Depends on: apply_scoping_display, DEFAULT_SCOPING_DISPLAY.
    Returns: nothing (side effect on the viewport).
    """
    apply_scoping_display(DEFAULT_SCOPING_DISPLAY)


NO_LEGEND_LABEL = "-- Automatic legend --"


def collect_legend_files():
    """
    Does: lists the legends available in LEGEND_FOLDER (.xml files).
    Depends on: os.path.isdir/os.listdir, LEGEND_FOLDER (00_constants.py).
    Returns: list, legend names without extension (str), sorted alphabetically, empty if the folder is absent.
    """
    if not os.path.isdir(LEGEND_FOLDER):
        return []
    return sorted(f[:-4] for f in os.listdir(LEGEND_FOLDER) if f.lower().endswith(".xml"))


def get_result_display_unit(result_obj, force_evaluate=True):
    """
    Does: extracts the displayed result's unit (e.g. "MPa") from the text of its Minimum/Maximum/Average property.
    Depends on: result_obj.EvaluateAllResults, result_obj.VisibleProperties (the Details panel).
    Returns: str, the detected unit (e.g. "MPa"), or None if unavailable.
    """
    # result_obj.Maximum.Unit turned out to be unreliable (unit missing/incorrect depending on the result type); the VisibleProperties text always contains the unit actually used.
    if force_evaluate:
        try:
            result_obj.EvaluateAllResults()
        except Exception:
            pass

    # "Minimum Occurs On"/"Maximum Occurs On" are deliberately excluded: their StringValue is a body name, not a numeric value followed by a unit.
    candidate_captions = ("Minimum", "Maximum", "Average", "Minimum Value", "Maximum Value", "Average Value")

    # Logging reserved for real calls (force_evaluate=True): otherwise simply opening the "..." window would flood the console without generating anything.
    try:
        for prop in result_obj.VisibleProperties:
            try:
                caption = prop.Caption
                if caption not in candidate_captions:
                    continue
                tokens = prop.StringValue.split()
                if len(tokens) >= 2:
                    if force_evaluate:
                        print "Unit detected for {} via '{}' ({}): {}".format(
                            result_obj.Name, caption, prop.StringValue, tokens[-1])
                    return tokens[-1]
            except Exception:
                pass
    except Exception as e:
        if force_evaluate:
            print "Unit unavailable for {}: {}".format(result_obj.Name, str(e))
        return None

    if force_evaluate:
        print "No unit detected for {} (no usable Minimum/Maximum/Average property).".format(result_obj.Name)
    return None


def apply_legend_if_exists(legend_name, result_obj):
    """
    Does: imports a legend by name in the unit of the relevant result and applies it to the viewport.
    Depends on: LEGEND_FOLDER, get_result_display_unit, ExtAPI.Graphics.ImportLegend, CurrentLegendSettings.
    Returns: nothing (side effect: changes the viewport legend, or does nothing if legend_name is None).
    """
    if not legend_name:
        return
    xml_path = os.path.join(LEGEND_FOLDER, legend_name + ".xml")
    if not os.path.exists(xml_path):
        print "Legend not found: " + xml_path
        return

    # ImportLegend/Reset compare the requested unit to that of the object CURRENTLY ACTIVE in the viewport, not to result_obj: without this explicit Activate(), the unit compared would stay the one from the previous row.
    try:
        result_obj.Activate()
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Unable to activate {} before applying the legend: {}".format(result_obj.Name, str(e))

    # Attempted systematically even if unit is None: lets the actual .NET error surface in the console instead of silently giving up.
    unit = get_result_display_unit(result_obj)
    print "Legend '{}' on {}: unit used for ImportLegend = {}".format(legend_name, result_obj.Name, unit)

    reset_legend()

    try:
        legend = ExtAPI.Graphics.ImportLegend(xml_path, unit)
        legend.CopyTo(Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings())
        print "Legend '{}' applied on {} (unit={}).".format(legend_name, result_obj.Name, unit)
    except Exception as e:
        print "Unable to apply legend '{}' on {} (unit={}): {}".format(
            legend_name, result_obj.Name, unit, str(e))


def reset_legend():
    """
    Does: resets the viewport's current legend to its automatic state.
    Depends on: Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings, ExtAPI.Graphics.Redraw.
    Returns: nothing (side effect: resets the displayed legend to automatic).
    """
    # Redraw() is essential: changing this property via script does not refresh the viewport on its own as long as no other event forces a redraw.
    try:
        Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings().Reset()
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Unable to reset legend: " + str(e)


# Available templates: 2, 3, 4, 6 and 8 steps; 5, 7 steps (and beyond 8) automatically fall back to "individual slides" mode.
MULTI_STEP_SLIDE_TEMPLATES = {
    2: {
        "layout_index": 3,
        "image_shape_indices": [3, 2],
    },
    3: {
        "layout_index": 4,
        "image_shape_indices": [3, 2, 8],
    },
    4: {
        "layout_index": 5,
        "image_shape_indices": [3, 2, 8, 9],
    },
    6: {
        "layout_index": 6,
        "image_shape_indices": [3, 2, 8, 9, 10, 11],
    },
    8: {
        "layout_index": 7,
        "image_shape_indices": [3, 2, 8, 9, 10, 11, 12, 13],
    },
}


def get_multi_step_template(step_count):
    """
    Does: retrieves the combined-slide template matching this exact number of steps.
    Depends on: MULTI_STEP_SLIDE_TEMPLATES.
    Returns: dict (layout_index/image_shape_indices), or None if no template exists for this number of steps.
    """
    return MULTI_STEP_SLIDE_TEMPLATES.get(step_count)


def get_step_count(analysis):
    """
    Does: reads the number of steps defined at the analysis level.
    Depends on: analysis.AnalysisSettings.NumberOfSteps.
    Returns: int, the number of steps, or 1 on error.
    """
    try:
        return int(analysis.AnalysisSettings.NumberOfSteps)
    except Exception as e:
        print "Step count unavailable: " + str(e)
        return 1


def _set_result_display_time(result_obj, display_time):
    """
    Does: repositions a result at a precise DisplayTime and re-evaluates it.
    Depends on: result_obj.DisplayTime/EvaluateAllResults/Evaluate, ExtAPI.Graphics.Redraw, SWF.Application.DoEvents.
    Returns: nothing (side effect: restores the result's original display).
    """
    # Only used to restore the original state afterwards: per-step captures go through evaluate_result_for_step (SetNumber), not through this function.
    result_obj.DisplayTime = display_time
    for method_name in ("EvaluateAllResults", "Evaluate"):
        method = getattr(result_obj, method_name, None)
        if method is None:
            continue
        try:
            method()
            break
        except Exception as e:
            print "Evaluation ({}) failed for {}: {}".format(method_name, result_obj.Name, str(e))
    try:
        ExtAPI.Graphics.Redraw()
        SWF.Application.DoEvents()
    except Exception as e:
        print "Redraw failed after re-evaluating {}: {}".format(result_obj.Name, str(e))


def evaluate_result_for_step(result_obj, step_number):
    """
    Does: positions a result on a precise set/step via SetNumber rather than DisplayTime.
    Depends on: result_obj.Activate/By/SetNumber/EvaluateAllResults, SetDriverStyle.ResultSet, ExtAPI.Graphics.Redraw.
    Returns: nothing (side effect: the result now displays this step).
    """
    # SetNumber navigates to a set already computed by the solver without a full re-evaluation: more reliable for chaining several steps than the old DisplayTime approach.
    result_obj.Activate()
    result_obj.By = SetDriverStyle.ResultSet
    result_obj.SetNumber = step_number
    result_obj.EvaluateAllResults()
    ExtAPI.Graphics.Redraw()


def export_result_image_for_step(result_obj, step_number):
    """
    Does: exports the image of a result for a precise set/step.
    Depends on: evaluate_result_for_step, export_current_view_image (02_image_export.py).
    Returns: str, the path of the generated PNG.
    """
    # Direct view export after Activate(), without a Figure snapshot (unlike export_object_image).
    evaluate_result_for_step(result_obj, step_number)
    return export_current_view_image("{}_step{}".format(result_obj.Name, step_number))


def add_multi_step_image_slide(presentation, template, title, image_paths):
    """
    Does: adds a combined slide (several step images) from a MULTI_STEP_SLIDE_TEMPLATES template.
    Depends on: presentation.SlideMaster.CustomLayouts, presentation.Slides.AddSlide.
    Returns: PPT.Slide, the created slide.
    """
    layout = presentation.SlideMaster.CustomLayouts[template["layout_index"]]
    slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout)
    slide.Shapes[1].TextFrame.TextRange.Text = title

    # Sorted by actual position (top->bottom, left->right), not by shape index, to respect the chronological order of the steps.
    placeholders = [slide.Shapes[idx] for idx in template["image_shape_indices"]]
    placeholders.sort(key=lambda ph: (ph.Top, ph.Left))

    for i in range(min(len(image_paths), len(placeholders))):
        ph = placeholders[i]
        slide.Shapes.AddPicture(image_paths[i], Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                 ph.Left, ph.Top, ph.Width, ph.Height)
    return slide


def capture_multi_result_cell_image(cfg, views, section_planes, section_labels):
    """
    Does: applies the graphic configuration of ONE cell (combined multi-result slide: view,
    section, legend, appearance, scale factor) and exports a single image of the chosen result - no
    notion of step here (unlike build_single_result_slide/build_step_based_result_slides),
    each cell carries a different result displayed in its current state.
    Depends on: apply_view_if_exists, apply_section_plane, apply_scale_factor, apply_contour_view,
    apply_legend_orientation, apply_scoping_display, apply_legend_if_exists, export_solution_image
    (02_image_export.py), disable_all_section_planes, reset_scale_factor/reset_contour_view/
    reset_legend_orientation/reset_scoping_display.
    Returns: str, the path of the exported image, or None on error.
    """
    obj = cfg.obj
    image_path = None
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        apply_scale_factor(cfg.deformation_scale_mode, cfg.scale_factor)
        apply_contour_view(cfg.contour_view)
        apply_legend_orientation(cfg.legend_orientation)
        apply_scoping_display(cfg.scoping_display)
        apply_legend_if_exists(cfg.legend_name, obj)
        image_path = export_solution_image(obj)
    except Exception as e:
        print "Unable to capture {}: {}".format(obj.Name, str(e))
    finally:
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
    return image_path


def build_multi_result_slide(report, template, cell_configs, views, section_planes, section_labels):
    """
    Does: builds ONE combined "different results" slide (a different result per configured
    cell, each with its own view/section/legend/appearance), from a
    MULTI_STEP_SLIDE_TEMPLATES template (same template family as the multi-step combined slides, but
    here each slot receives a different result rather than the same result at a different step).
    Depends on: capture_multi_result_cell_image, add_multi_step_image_slide.
    Returns: nothing (side effect: adds a slide to report.presentation, or does nothing if no cell is configured).
    """
    image_paths = []
    titles = []
    for cfg in cell_configs:
        if cfg is None:
            continue
        image_path = capture_multi_result_cell_image(cfg, views, section_planes, section_labels)
        if image_path:
            image_paths.append(image_path)
            titles.append(cfg.obj.Name)

    if not image_paths:
        print "No cell configured: combined multi-result slide not generated."
        return

    title = "Combined results: " + ", ".join(titles)
    add_multi_step_image_slide(report.presentation, template, title, image_paths)
    print "Combined multi-result slide added ({} results).".format(len(image_paths))


def build_step_based_result_slides(report, cfg, obj, subtitle, analysis):
    """
    Does: builds the slide(s) for a result with a step selection (one combined slide if possible, otherwise one per step).
    Depends on: get_multi_step_template, export_result_image_for_step, export_result_tabular_data, report.add_image_table_slide.
    Returns: nothing (side effect: adds one or more slides to the report).
    """
    # analysis is no longer used (captures via SetNumber, not via a time computed from the analysis); kept in the signature so as not to break callers.
    steps = cfg.selected_steps
    template = get_multi_step_template(len(steps)) if cfg.step_display_mode == "combined" else None
    original_display_time = obj.DisplayTime
    original_by = obj.By

    display_name = obj.Name + analysis_suffix(cfg)

    try:
        if template:
            image_paths = [export_result_image_for_step(obj, step) for step in steps]
            title = "{} - {} steps".format(display_name, len(steps))
            add_multi_step_image_slide(report.presentation, template, title, image_paths)
            print "Combined slide ({} steps) added for {}.".format(len(steps), display_name)
            return

        csv_path = None
        try:
            csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, obj)
        except Exception as e:
            print "Unable to export CSV for {}: {}".format(obj.Name, str(e))

        for step in steps:
            img_path = None
            try:
                img_path = export_result_image_for_step(obj, step)
            except Exception as e:
                print "Unable to export image for {} (step {}): {}".format(obj.Name, step, str(e))
            title = "{} - Step {}".format(display_name, step)
            report.add_image_table_slide(title, subtitle, img_path=img_path, csv_path=csv_path)
    finally:
        # Restoration mandatory even on error: otherwise the object stays frozen on the last processed step and throws off the legend of the following slides.
        obj.By = original_by
        _set_result_display_time(obj, original_display_time)


def flatten_results(objects):
    """
    Does: recursively unfolds grouping folders (e.g. "Group Similar Children") into their leaf objects.
    Depends on: nothing (recursively iterates over obj.Children).
    Returns: list, the leaf objects exportable to image/CSV.
    """
    leaves = []
    for obj in objects:
        try:
            children = obj.Children
        except Exception:
            children = None
        if children is not None and len(children) > 0:
            leaves.extend(flatten_results(list(children)))
        else:
            leaves.append(obj)
    return leaves


def collect_boundary_conditions(analysis=None):
    """
    Does: lists the Boundary Conditions in the model, limited to one analysis if given.
    Depends on: ExtAPI.DataModel.GetObjectsByType(GenericBoundaryCondition), _is_descendant_of.
    Returns: list, the Boundary Condition objects (whole project if analysis is None).
    """
    all_bcs = list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.GenericBoundaryCondition))
    if analysis is None:
        return all_bcs
    return [bc for bc in all_bcs if _is_descendant_of(bc, analysis)]


def collect_boundary_conditions_multi(analyses):
    """
    Does: lists the Boundary Conditions of all given analyses.
    Depends on: collect_boundary_conditions.
    Returns: list of tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for bc in collect_boundary_conditions(analysis):
            pairs.append((bc, analysis))
    return pairs


def collect_analyses():
    """
    Does: lists the project's analyses usable by the generator (valid Analysis Settings).
    Depends on: ExtAPI.DataModel.AnalysisList, analysis.AnalysisSettings.NumberOfSteps.
    Returns: list, the project's analysis objects - excludes post-processing addins (e.g. FEMFAT)
    whose AnalysisSettings is None: they have neither steps nor classic solution results, and
    used to crash the whole generation (settings.NumberOfSteps on a None object) as soon as they
    were selected in a GUI list.
    """
    analyses = []
    for analysis in ExtAPI.DataModel.AnalysisList:
        try:
            analysis.AnalysisSettings.NumberOfSteps
        except Exception:
            print "Analysis skipped (Analysis Settings unavailable, e.g. FEMFAT addin): " + str(analysis.Name)
            continue
        analyses.append(analysis)
    return analyses


def collect_bolt_pretensions(analysis=None):
    """
    Does: lists the Bolt Pretensions in the model, limited to one analysis if given.
    Depends on: ExtAPI.DataModel.GetObjectsByType(BoltPretension), _is_descendant_of.
    Returns: list, the Bolt Pretension objects (whole project if analysis is None).
    """
    all_bolt_pretensions = list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltPretension))
    if analysis is None:
        return all_bolt_pretensions
    return [bp for bp in all_bolt_pretensions if _is_descendant_of(bp, analysis)]


def collect_bolt_pretensions_multi(analyses):
    """
    Does: lists the Bolt Pretensions of all given analyses.
    Depends on: collect_bolt_pretensions.
    Returns: list of tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for bp in collect_bolt_pretensions(analysis):
            pairs.append((bp, analysis))
    return pairs


def _is_descendant_of(obj, ancestor):
    """
    Does: checks whether obj is a descendant (direct or indirect) of ancestor in the Mechanical tree.
    Depends on: obj.Parent (walking up the tree).
    Returns: bool, True if ancestor is indeed a parent of obj.
    """
    # Distinguishes a Contact Tool from the Connections branch (without step) from its namesake in Solution (with step): same .NET category, only the position in the tree differs.
    node = getattr(obj, "Parent", None)
    while node is not None:
        if node == ancestor:
            return True
        node = getattr(node, "Parent", None)
    return False


def collect_contact_tool_results(analysis):
    """
    Does: lists the results of the Contact Tool folders from the Solution branch (with steps) for an analysis.
    Depends on: ExtAPI.DataModel.GetObjectsByType(ContactTool), _is_descendant_of, flatten_results.
    Returns: list, the exportable result objects specific to the Solution branch.
    """
    # ContactTool also exists in Connections (without step, same children names): filtering by branch avoids mixing the two lists.
    tools = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactTool)
    children = []
    for tool in tools:
        if tool.Children and _is_descendant_of(tool, analysis.Solution):
            children.extend(list(tool.Children))
    return flatten_results(children)


def collect_contact_tool_results_multi(analyses):
    """
    Does: lists the Contact Tool results (Solution branch) of all given analyses.
    Depends on: collect_contact_tool_results.
    Returns: list of tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for obj in collect_contact_tool_results(analysis):
            pairs.append((obj, analysis))
    return pairs


def collect_connection_contact_tool_results(analysis):
    """
    Does: lists the results of the Contact Tool folders from the Connections branch (without step) for an analysis.
    Depends on: ExtAPI.DataModel.GetObjectsByType(ContactTool), _is_descendant_of, flatten_results.
    Returns: list, the exportable result objects specific to the Connections branch.
    """
    tools = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactTool)
    children = []
    for tool in tools:
        if tool.Children and not _is_descendant_of(tool, analysis.Solution):
            children.extend(list(tool.Children))
    return flatten_results(children)


def collect_bolt_tool_results(analysis=None):
    """
    Does: lists the results of the Bolt Tool folders under Solution, limited to one analysis if given.
    Depends on: ExtAPI.DataModel.GetObjectsByType(BoltTool), _is_descendant_of, flatten_results.
    Returns: list, the exportable result objects (whole project if analysis is None).
    """
    tools = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltTool)
    children = []
    for tool in tools:
        if not tool.Children:
            continue
        if analysis is not None and not _is_descendant_of(tool, analysis.Solution):
            continue
        children.extend(list(tool.Children))
    return flatten_results(children)


def collect_bolt_tool_results_multi(analyses):
    """
    Does: lists the Bolt Tool results of all given analyses.
    Depends on: collect_bolt_tool_results.
    Returns: list of tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for obj in collect_bolt_tool_results(analysis):
            pairs.append((obj, analysis))
    return pairs


def collect_all_results(analysis):
    """
    Does: lists the "simple" Solution results (Deformation, Stress, Probe...), excluding Solution Information/Contact Tool/Bolt Tool.
    Depends on: analysis.Solution.Children, ExtAPI.DataModel.GetObjectsByType(ContactTool/BoltTool), flatten_results.
    Returns: list, the exportable result objects.
    """
    excluded_categories = [DataModelObjectCategory.ContactTool, DataModelObjectCategory.BoltTool]

    # Exclusion by identity in addition to category: DataModelObjectCategory can fail silently (category=None) on the Contact/Bolt Tool folder itself, which would let it pass the filter and duplicate its children with the dedicated separate list.
    already_handled = list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactTool))
    already_handled += list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltTool))

    solution_children = analysis.Solution.Children

    candidates = []
    if solution_children:
        for i in range(1, len(solution_children)):  # index 0 = Solution Information
            child = solution_children[i]
            if child in already_handled:
                continue
            try:
                category = child.DataModelObjectCategory
            except Exception:
                category = None
            if category in excluded_categories:
                continue
            candidates.append(child)

    return flatten_results(candidates)


def collect_all_results_multi(analyses):
    """
    Does: lists the simple results of all given analyses.
    Depends on: collect_all_results.
    Returns: list of tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for obj in collect_all_results(analysis):
            pairs.append((obj, analysis))
    return pairs


def build_single_bc_slide(report, cfg, views, section_planes, section_labels):
    """
    Does: builds the slide for ONE Boundary Condition with its configured view/section/legend/appearance.
    Depends on: apply_view_if_exists, apply_section_plane, apply_scale_factor, apply_legend_if_exists, apply_contour_view, apply_legend_orientation, apply_scoping_display, export_object_image, export_bc_tabular_data, report.add_image_table_slide.
    Returns: nothing (side effect: adds a slide to the report).
    """
    bc = cfg.obj
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        apply_scale_factor(cfg.deformation_scale_mode, cfg.scale_factor)
        apply_contour_view(cfg.contour_view)
        apply_legend_orientation(cfg.legend_orientation)
        apply_scoping_display(cfg.scoping_display)
        apply_legend_if_exists(cfg.legend_name, bc)
        img_path = export_object_image(bc, bc.Name)
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
        csv_path = export_bc_tabular_data(CSV_EXPORT_FOLDER, bc)
        report.add_image_table_slide(bc.Name + analysis_suffix(cfg), "-- Boundary Conditions --",
                                      img_path=img_path, csv_path=csv_path)
    except Exception as e:
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
        print "Unable to build BC slide for {}: {}".format(bc.Name, str(e))


def build_bc_slides(report, row_configs, views, section_planes, section_labels):
    """
    Does: adds a slide for each selected Boundary Condition.
    Depends on: build_single_bc_slide.
    Returns: nothing (side effect: adds a slide per row to the report).
    """
    for cfg in row_configs:
        build_single_bc_slide(report, cfg, views, section_planes, section_labels)


def build_single_bp_slide(report, cfg, views, section_planes, section_labels):
    """
    Does: builds the slide for ONE Bolt Pretension with its configured view/section/legend/appearance.
    Depends on: apply_view_if_exists, apply_section_plane, apply_scale_factor, apply_legend_if_exists, apply_contour_view, apply_legend_orientation, apply_scoping_display, export_object_image, export_bp_tabular_data, report.add_image_table_slide.
    Returns: nothing (side effect: adds a slide to the report).
    """
    bp = cfg.obj
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        apply_scale_factor(cfg.deformation_scale_mode, cfg.scale_factor)
        apply_contour_view(cfg.contour_view)
        apply_legend_orientation(cfg.legend_orientation)
        apply_scoping_display(cfg.scoping_display)
        apply_legend_if_exists(cfg.legend_name, bp)
        img_path = export_object_image(bp, bp.Name)
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
        csv_path = export_bp_tabular_data(CSV_EXPORT_FOLDER, bp)
        report.add_image_table_slide(bp.Name + analysis_suffix(cfg), "-- Bolt Pretension --",
                                      img_path=img_path, csv_path=csv_path)
    except Exception as e:
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
        print "Unable to build Bolt Pretension slide for {}: {}".format(bp.Name, str(e))


def build_bp_slides(report, row_configs, views, section_planes, section_labels):
    """
    Does: adds a slide for each selected Bolt Pretension.
    Depends on: build_single_bp_slide.
    Returns: nothing (side effect: adds a slide per row to the report).
    """
    for cfg in row_configs:
        build_single_bp_slide(report, cfg, views, section_planes, section_labels)


def build_single_result_slide(report, cfg, subtitle, views, section_planes, section_labels, analysis):
    """
    Does: builds the slide for ONE result object, with its view/section/legend/appearance and optional step selection.
    Depends on: apply_view_if_exists, apply_section_plane, apply_scale_factor, apply_legend_if_exists, apply_contour_view, apply_legend_orientation, apply_scoping_display, build_step_based_result_slides, export_object_image, export_result_tabular_data.
    Returns: nothing (side effect: adds one or more slides to the report).
    """
    obj = cfg.obj
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        apply_scale_factor(cfg.deformation_scale_mode, cfg.scale_factor)
        apply_contour_view(cfg.contour_view)
        apply_legend_orientation(cfg.legend_orientation)
        apply_scoping_display(cfg.scoping_display)
        apply_legend_if_exists(cfg.legend_name, obj)

        if cfg.selected_steps:
            build_step_based_result_slides(report, cfg, obj, subtitle, analysis)
        else:
            img_path = None
            try:
                img_path = export_object_image(obj, obj.Name)
            except Exception as e:
                print "Unable to export image for {}: {}".format(obj.Name, str(e))

            csv_path = None
            try:
                csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, obj)
            except Exception as e:
                print "Unable to export CSV for {}: {}".format(obj.Name, str(e))

            if img_path or csv_path:
                report.add_image_table_slide(obj.Name + analysis_suffix(cfg), subtitle,
                                              img_path=img_path, csv_path=csv_path)
            else:
                print "No exportable data for " + obj.Name + ": slide skipped."
    except Exception as e:
        print "Unable to build result slide for {}: {}".format(obj.Name, str(e))
    finally:
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()


def build_result_slides(report, row_configs, subtitle, views, section_planes, section_labels, analysis):
    """
    Does: adds a slide for each selected result object (Contact Tool, Bolt Tool, or general results).
    Depends on: build_single_result_slide.
    Returns: nothing (side effect: adds a slide per row to the report).
    """
    for cfg in row_configs:
        build_single_result_slide(report, cfg, subtitle, views, section_planes, section_labels, analysis)


DEFAULT_CONTEXT_OPACITY_PERCENT = 25


class GeometryPartRowConfig(object):
    """
    Display configuration for a "simple geometry" slide (one isolated, opaque part, within the transparent context of the assembly).
    """

    def __init__(self, body):
        """
        Does: initializes an isolated part's (simple geometry) configuration with its default values.
        Depends on: DEFAULT_CONTEXT_OPACITY_PERCENT.
        Returns: nothing (constructor).
        """
        self.obj = body
        self.view_name = None
        self.section_name = None
        self.context_opacity_percent = DEFAULT_CONTEXT_OPACITY_PERCENT
        self.configured = False


def build_geometry_row_display_name(row_config):
    """
    Does: builds the text shown in the list for a part (simple geometry).
    Depends on: nothing (reads row_config.obj/view_name/section_name/context_opacity_percent).
    Returns: str, the part's name followed by the chosen settings separated by " | ".
    """
    parts = [row_config.obj.Name]
    if row_config.view_name:
        parts.append("view=" + row_config.view_name)
    if row_config.section_name:
        parts.append("section=" + row_config.section_name)
    parts.append("context={}%".format(row_config.context_opacity_percent))
    return " | ".join(parts)


def collect_bodies():
    """
    Does: lists all the bodies (Body) in the model.
    Depends on: ExtAPI.DataModel.Project.Model.GetChildren(DataModelObjectCategory.Body).
    Returns: list, the model's Body objects.
    """
    return list(ExtAPI.DataModel.Project.Model.GetChildren(DataModelObjectCategory.Body, True))


def isolate_body_by_transparency(target_body, all_bodies, context_opacity_percent):
    """
    Does: makes one part fully opaque and the others semi-transparent at the given percentage.
    Depends on: Body.Transparency, Transaction (Ansys.ACT.Mechanical), ExtAPI.Graphics.Redraw.
    Returns: nothing (side effect: changes the transparency of all bodies in the model).
    """
    # Body.Transparency ranges from 0.0 (transparent) to 1.0 (opaque), despite its name suggesting the opposite.
    context_value = max(0.0, min(1.0, context_opacity_percent / 100.0))
    target_id = target_body.ObjectId
    # Transaction(True) defers the refresh during the loop; the explicit Redraw() afterwards is still necessary for the change to be visible before the image capture.
    with Transaction(True):
        for body in all_bodies:
            try:
                body.Transparency = 1.0 if body.ObjectId == target_id else context_value
            except Exception:
                pass
    try:
        ExtAPI.Graphics.Redraw()
    except Exception:
        pass


def reset_all_bodies_transparency(all_bodies):
    """
    Does: restores all parts to normal opacity (100%).
    Depends on: Body.Transparency, Transaction (Ansys.ACT.Mechanical), ExtAPI.Graphics.Redraw.
    Returns: nothing (side effect: restores the opacity of all bodies in the model).
    """
    with Transaction(True):
        for body in all_bodies:
            try:
                body.Transparency = 1.0
            except Exception:
                pass
    try:
        ExtAPI.Graphics.Redraw()
    except Exception:
        pass


def export_geometry_part_image(body, all_bodies, context_opacity_percent):
    """
    Does: isolates a part (opaque) within the transparent context of the assembly, then exports its image.
    Depends on: isolate_body_by_transparency, geometry.AddFigure, export_current_view_image, reset_all_bodies_transparency.
    Returns: str, the path of the generated PNG.
    """
    isolate_body_by_transparency(body, all_bodies, context_opacity_percent)
    geometry = ExtAPI.DataModel.Project.Model.Geometry
    figure = geometry.AddFigure()
    figure.Activate()
    # No SetFit(): the camera framing (view chosen via apply_view_if_exists, or the current
    # manual position) is left as is, at the user's responsibility.
    image_path = export_current_view_image("Geometry_" + safe_file_name(body.Name))
    reset_all_bodies_transparency(all_bodies)
    return image_path


def build_single_geometry_part_slide(report, cfg, all_bodies, views, section_planes, section_labels):
    """
    Does: builds the "simple geometry" slide (one image, no table) for ONE isolated part.
    Depends on: apply_view_if_exists, apply_section_plane, export_geometry_part_image, report.add_image_table_slide.
    Returns: nothing (side effect: adds a slide to the report).
    """
    body = cfg.obj
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        img_path = export_geometry_part_image(body, all_bodies, cfg.context_opacity_percent)
        report.add_image_table_slide(body.Name, "-- Geometry --", img_path=img_path, csv_path=None)
    except Exception as e:
        print "Unable to build geometry slide for {}: {}".format(body.Name, str(e))
    finally:
        disable_all_section_planes(section_planes)


def build_geometry_part_slides(report, row_configs, all_bodies, views, section_planes, section_labels):
    """
    Does: adds a "simple geometry" slide for each selected part.
    Depends on: build_single_geometry_part_slide.
    Returns: nothing (side effect: adds a slide per part to the report).
    """
    for cfg in row_configs:
        build_single_geometry_part_slide(report, cfg, all_bodies, views, section_planes, section_labels)


class MeshPartRowConfig(object):
    """
    Selection row for mesh by isolated part: the body and an optional view (no section/opacity, isolation via full hiding).
    """

    def __init__(self, body):
        """
        Does: initializes an isolated part's (mesh) configuration with its default values.
        Depends on: nothing (simple assignments).
        Returns: nothing (constructor).
        """
        self.obj = body
        self.view_name = None
        self.configured = False  # becomes True once the "..." button has been confirmed (OK)


def build_mesh_part_row_display_name(row_config):
    """
    Does: builds the text shown in the list for a part (mesh by isolated part).
    Depends on: nothing (reads row_config.obj/view_name).
    Returns: str, the part's name followed by the chosen view, separated by " | ".
    """
    parts = [row_config.obj.Name]
    if row_config.view_name:
        parts.append("view=" + row_config.view_name)
    return " | ".join(parts)


def show_only_body(target_body, all_bodies):
    """
    Does: hides all parts except target_body.
    Depends on: Body.Visible, Transaction (Ansys.ACT.Mechanical).
    Returns: nothing (side effect: changes the visibility of all bodies in the model).
    """
    # Transaction(True) defers the refresh during the loop; the render is captured later anyway, during the image export.
    target_id = target_body.ObjectId
    with Transaction(True):
        for body in all_bodies:
            try:
                body.Visible = (body.ObjectId == target_id)
            except Exception:
                pass


def show_all_bodies(all_bodies):
    """
    Does: makes all parts visible.
    Depends on: Body.Visible, Transaction (Ansys.ACT.Mechanical).
    Returns: nothing (side effect: restores the visibility of all bodies in the model).
    """
    with Transaction(True):
        for body in all_bodies:
            try:
                body.Visible = True
            except Exception:
                pass


def get_body_mesh_counts(body):
    """
    Does: retrieves the node/element count of ONE body.
    Depends on: body.VisibleProperties (the Details panel).
    Returns: tuple (node_count, element_count), each None if unavailable.
    """
    # Read via VisibleProperties (Details panel): MeshRegionById(body.ObjectId) turned out to be misattributed (values at 0 or inconsistent between parts).
    node_count = None
    element_count = None
    try:
        for prop in body.VisibleProperties:
            if prop.Name == "Nodes":
                node_count = prop.StringValue
            elif prop.Name == "Elements":
                element_count = prop.StringValue
    except Exception as e:
        print "Unable to count nodes/elements for {}: {}".format(body.Name, str(e))
    return node_count, element_count


def export_body_mesh_image(body, all_bodies, image_name):
    """
    Does: isolates a part (other bodies hidden) and exports an image of its mesh.
    Depends on: show_only_body, Model.Mesh.AddFigure, export_current_view_image, show_all_bodies.
    Returns: str, the path of the generated PNG.
    """
    # ExtAPI.Graphics.ViewOptions.ShowMesh must already be forced to True by the caller (build_mesh_part_slides): a single force/reset for a whole group of captures.
    show_only_body(body, all_bodies)
    mesh = ExtAPI.DataModel.Project.Model.Mesh
    figure = mesh.AddFigure()
    figure.Activate()
    # SetFit() is necessary here (unlike the rest of the project, see README): show_only_body()
    # completely hides the other bodies (Visible=False, no transparent context like for the
    # geometry), so the only content in the viewport is the targeted part and SetFit() cannot overwrite
    # another useful view -- without this call, the camera keeps the framing of the full assembly and a
    # small isolated part (e.g. a bolt) appears tiny in the exported image.
    ExtAPI.Graphics.Camera.SetFit()
    image_path = export_current_view_image(image_name)
    show_all_bodies(all_bodies)
    return image_path


def export_body_mesh_summary_csv(directory, body):
    """
    Does: exports a minimal mesh statistics table (ElementSize, Nodes, Elements) for ONE part.
    Depends on: get_body_mesh_counts, Model.Mesh, _format_element_size (01_data_export.py) for ElementSize, get_unique_file_path/to_csv_cell (00_constants.py), the csv module.
    Returns: str, the path of the generated CSV.
    """
    mesh = Model.Mesh
    node_count, element_count = get_body_mesh_counts(body)

    rows = [
        ["ElementSize", _format_element_size(mesh)],
        ["Nodes", node_count],
        ["Elements", element_count],
    ]

    filepath = get_unique_file_path(directory, "Mesh_" + safe_file_name(body.Name), ".csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Property", "Value"])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "CSV export complete: " + filepath
    return filepath


def add_mesh_multi_image_slide(report, image_paths, csv_paths):
    """
    Does: adds a LAYOUT_MESH_MULTI slide (up to 4 images + 4 tables) for a group of parts.
    Depends on: LAYOUT_MESH_MULTI, MESH_MULTI_IMAGE_SHAPE_INDICES, MESH_MULTI_TABLE_SHAPE_INDICES, report.add_csv_table.
    Returns: PPT.Slide, the created slide.
    """
    layout = report.presentation.SlideMaster.CustomLayouts[LAYOUT_MESH_MULTI]
    slide = report.presentation.Slides.AddSlide(report.presentation.Slides.Count + 1, layout)

    slide.Shapes[1].TextFrame.TextRange.Text = "Mesh Details"

    # Sorted by actual position (top->bottom, left->right), not by shape index, same precaution as add_multi_step_image_slide.
    image_placeholders = [slide.Shapes[i] for i in MESH_MULTI_IMAGE_SHAPE_INDICES]
    image_placeholders.sort(key=lambda ph: (ph.Top, ph.Left))

    table_placeholders = [slide.Shapes[i] for i in MESH_MULTI_TABLE_SHAPE_INDICES]
    table_placeholders.sort(key=lambda ph: (ph.Top, ph.Left))

    for i in range(min(len(image_paths), len(image_placeholders))):
        ph = image_placeholders[i]
        if image_paths[i]:
            slide.Shapes.AddPicture(image_paths[i], Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                     ph.Left, ph.Top, ph.Width, ph.Height)

    for i in range(min(len(csv_paths), len(table_placeholders))):
        ph = table_placeholders[i]
        if csv_paths[i]:
            try:
                report.add_csv_table(slide, csv_paths[i], ph.Left, ph.Top, ph.Width)
            except Exception as e:
                print "Unable to insert table ({}): {}".format(csv_paths[i], str(e))

    return slide


def build_mesh_part_slides(report, row_configs, all_bodies, views):
    """
    Does: adds one or more "mesh by isolated part" slides, grouping parts by MAX_MESH_MULTI_BODIES.
    Depends on: apply_view_if_exists, export_body_mesh_image, export_body_mesh_summary_csv, add_mesh_multi_image_slide.
    Returns: nothing (side effect: adds one or more slides to the report).
    """
    for start in range(0, len(row_configs), MAX_MESH_MULTI_BODIES):
        chunk = row_configs[start:start + MAX_MESH_MULTI_BODIES]

        ExtAPI.Graphics.ViewOptions.ShowMesh = True
        image_paths = []
        try:
            for cfg in chunk:
                body = cfg.obj
                try:
                    apply_view_if_exists(cfg.view_name, views)
                    image_paths.append(export_body_mesh_image(body, all_bodies, "Mesh_" + safe_file_name(body.Name)))
                except Exception as e:
                    print "Unable to export image for {}: {}".format(body.Name, str(e))
                    image_paths.append(None)
        finally:
            ExtAPI.Graphics.ViewOptions.ShowMesh = False

        csv_paths = []
        for cfg in chunk:
            body = cfg.obj
            try:
                csv_paths.append(export_body_mesh_summary_csv(CSV_EXPORT_FOLDER, body))
            except Exception as e:
                print "Unable to export CSV for {}: {}".format(body.Name, str(e))
                csv_paths.append(None)

        add_mesh_multi_image_slide(report, image_paths, csv_paths)
        print "Mesh multi-image slide added ({} part(s)).".format(len(chunk))


class ContactRowConfig(object):
    """
    Selection row for the Contact summary slide: just the contact, nothing to configure (all checked rows share ONE single slide).
    """

    def __init__(self, contact_region):
        """
        Does: initializes a Contact summary row's configuration (nothing to configure).
        Depends on: nothing (simple assignments).
        Returns: nothing (constructor).
        """
        self.obj = contact_region
        self.configured = True  # no "..." button for this category: always "ready"


def build_contact_row_display_name(row_config):
    """
    Does: builds the text shown in the list for a Contact Region.
    Depends on: nothing (reads row_config.obj.Name).
    Returns: str, the contact's name.
    """
    return row_config.obj.Name


def collect_contact_regions():
    """
    Does: lists all Contact Regions in the model.
    Depends on: ExtAPI.DataModel.GetObjectsByType(ContactRegion).
    Returns: list, the Contact Region objects (Connections folder).
    """
    return list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactRegion))


def build_contact_summary_slide(report, row_configs):
    """
    Does: adds THE contacts summary slide, limited to the selected contacts.
    Depends on: export_contacts_summary_csv (01_data_export.py), report.add_table_slide.
    Returns: nothing (side effect: adds a slide to the report).
    """
    contact_list = [cfg.obj for cfg in row_configs]
    csv_path = export_contacts_summary_csv(CSV_EXPORT_FOLDER, contact_list)
    report.add_table_slide("Contacts summary", "-- Contact --", csv_path)


CURVE_COLOR_OPTIONS = [
    ("Automatic", None),
    ("Red", Color.IndianRed),
    ("Blue", Color.SteelBlue),
    ("Green", Color.SeaGreen),
    ("Orange", Color.DarkOrange),
    ("Purple", Color.MediumPurple),
    ("Black", Color.Black),
    ("Gray", Color.Gray),
]


def curve_color_label(color):
    """
    Does: finds the label shown for a curve color.
    Depends on: CURVE_COLOR_OPTIONS.
    Returns: str, the matching label, or "Automatic" if color is None or unknown.
    """
    if color is not None:
        for label, option_color in CURVE_COLOR_OPTIONS:
            if option_color is not None and option_color == color:
                return label
    return CURVE_COLOR_OPTIONS[0][0]


def curve_color_from_label(label):
    """
    Does: finds the color associated with a label from CURVE_COLOR_OPTIONS.
    Depends on: CURVE_COLOR_OPTIONS.
    Returns: Color or None, the matching color (None for "Automatic" or an unknown label).
    """
    for option_label, option_color in CURVE_COLOR_OPTIONS:
        if option_label == label:
            return option_color
    return None


class SolutionInfoRowConfig(object):
    """
    Display configuration for a Solution Information tracker: the object and the chart parameters (title, axes, color), None = inferred from the CSV.
    """

    def __init__(self, tracker, analysis=None):
        """
        Does: initializes a Solution Information tracker's configuration with its default values.
        Depends on: nothing (simple assignments).
        Returns: nothing (constructor).
        """
        self.obj = tracker
        self.analysis = analysis  # see SlideRowConfig.analysis / analysis_suffix
        self.chart_title = None
        self.x_axis_label = None
        self.y_axis_label = None
        self.curve_color = None
        self.configured = False


def build_solution_info_row_display_name(row_config):
    """
    Does: builds the text shown in the list for a Solution Information tracker.
    Depends on: analysis_suffix, row_config (chart_title, x_axis_label, y_axis_label, curve_color).
    Returns: str, the tracker's name followed by the chosen chart settings, separated by " | ".
    """
    parts = [row_config.obj.Name + analysis_suffix(row_config)]
    if row_config.chart_title:
        parts.append("title=" + row_config.chart_title)
    if row_config.x_axis_label:
        parts.append("x=" + row_config.x_axis_label)
    if row_config.y_axis_label:
        parts.append("y=" + row_config.y_axis_label)
    if row_config.curve_color is not None:
        parts.append("color=" + curve_color_label(row_config.curve_color))
    return " | ".join(parts)


def collect_solution_information_trackers(analysis):
    """
    Does: lists the trackers (children) of Solution Information for an analysis.
    Depends on: analysis.Solution.Children[0] (1st child of the Solution branch).
    Returns: list, the tracker objects, empty on error.
    """
    try:
        solution_information = analysis.Solution.Children[0]
        children = solution_information.Children
        return list(children) if children else []
    except Exception as e:
        print "Solution Information unavailable: " + str(e)
        return []


def collect_solution_information_trackers_multi(analyses):
    """
    Does: lists the Solution Information trackers of all given analyses.
    Depends on: collect_solution_information_trackers.
    Returns: list of tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for tracker in collect_solution_information_trackers(analysis):
            pairs.append((tracker, analysis))
    return pairs


def build_single_solution_info_slide(report, cfg):
    """
    Does: builds the slide for ONE Solution Information tracker, with its optional chart parameters.
    Depends on: export_result_tabular_data, export_chart_image_from_csv (02_image_export.py), get_scoped_contact_region_name (04_slides.py), report.add_image_table_slide.
    Returns: nothing (side effect: adds a slide to the report, or nothing if no exportable data).
    """
    tracker = cfg.obj
    csv_path = None
    try:
        csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, tracker)
    except Exception as e:
        print "Unable to export CSV for {}: {}".format(tracker.Name, str(e))

    img_path = None
    if csv_path:
        try:
            img_path = export_chart_image_from_csv(
                csv_path, tracker.Name, chart_title=cfg.chart_title, x_axis_label=cfg.x_axis_label,
                y_axis_label=cfg.y_axis_label, curve_color=cfg.curve_color
            )
        except Exception as e:
            print "Unable to build chart for {}: {}".format(tracker.Name, str(e))

    if img_path or csv_path:
        contact_region_name = get_scoped_contact_region_name(tracker)
        title = tracker.Name + analysis_suffix(cfg)
        if contact_region_name:
            title = "{} - {}".format(title, contact_region_name)
        report.add_image_table_slide(title, "-- Solution Information --", img_path=img_path, csv_path=csv_path)
    else:
        print "No exportable data for " + tracker.Name + ": slide skipped."


def build_solution_info_slides(report, row_configs):
    """
    Does: adds a slide for each selected Solution Information tracker.
    Depends on: build_single_solution_info_slide.
    Returns: nothing (side effect: adds a slide per tracker to the report).
    """
    for cfg in row_configs:
        build_single_solution_info_slide(report, cfg)


def export_mesh_summary_csv(directory):
    """
    Does: exports a minimal mesh statistics table (ElementSize, Nodes, Elements).
    Depends on: Model.Mesh, the csv module, _format_element_size (01_data_export.py) for ElementSize.
    Returns: str, the path of the generated CSV.
    """
    mesh = Model.Mesh
    rows = [
        ["ElementSize", _format_element_size(mesh)],
        ["Nodes", mesh.Nodes],
        ["Elements", mesh.Elements],
    ]

    filepath = os.path.join(directory, "mesh_summary.csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Property", "Value"])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "CSV export complete: " + filepath
    return filepath


def build_mesh_slide(report, use_full_table):
    """
    Does: adds the mesh slide (view + full table or summary depending on use_full_table).
    Depends on: export_mesh_image, export_mesh_report_csv (01_data_export.py), export_mesh_summary_csv, report.add_image_table_slide.
    Returns: nothing (side effect: adds a slide to the report).
    """
    img_path = export_mesh_image()
    if use_full_table:
        csv_path = export_mesh_report_csv(CSV_EXPORT_FOLDER)
    else:
        csv_path = export_mesh_summary_csv(CSV_EXPORT_FOLDER)
    report.add_image_table_slide("Mesh and mesh details", "-- Mesh --", img_path=img_path, csv_path=csv_path)


class AnalysisContextRowConfig(object):
    """
    Selection row for a context slide (Analysis Parameters): the analysis itself and an optional view (View Manager).
    """

    def __init__(self, analysis):
        """
        Does: initializes an Analysis Context row's configuration (optional view).
        Depends on: nothing (simple assignments).
        Returns: nothing (constructor).
        """
        self.obj = analysis
        self.view_name = None
        self.configured = False  # becomes True once the "..." button has been confirmed (OK)


def build_analysis_context_row_display_name(row_config):
    """
    Does: builds the text shown in the list for an analysis.
    Depends on: nothing (reads row_config.obj.Name/view_name).
    Returns: str, the analysis' name, followed by the chosen view if defined.
    """
    parts = [row_config.obj.Name]
    if row_config.view_name:
        parts.append("view=" + row_config.view_name)
    return " | ".join(parts)


def build_analysis_context_slides(report, row_configs, views):
    """
    Does: adds a context slide (Analysis Parameters) for each selected analysis, with its configured view.
    Depends on: apply_view_if_exists, create_analysis_parameters_slide (04_slides.py).
    Returns: nothing (side effect: adds a slide per analysis to the report).
    """
    for cfg in row_configs:
        apply_view_if_exists(cfg.view_name, views)
        create_analysis_parameters_slide(report, cfg.obj)
