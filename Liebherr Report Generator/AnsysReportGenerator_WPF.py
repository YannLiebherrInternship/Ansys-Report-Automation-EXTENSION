# AnsysReportGenerator_WPF.py : WPF entry point for the report generation application. Loads modules 00_constants.py to 05_interactive_slides.py (same folder as this script) via execfile(), then builds the window from AnsysReportGenerator_WPF.xaml.

import os
import shutil
import sys
import xml.etree.ElementTree as ET

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")
clr.AddReference("System")

from System.Diagnostics import Process

# System.Windows.Forms/Color remain necessary here: 05_interactive_slides.py calls
# SWF.Application.DoEvents() (_set_result_display_time) and uses named Color values for
# CURVE_COLOR_OPTIONS, even though THIS script builds its interface in WPF.
import System.Windows.Forms as SWF
from System.Drawing import Color

from System.IO import StreamReader
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader
from System.Windows import (
    Thickness, CornerRadius, GridLength, GridUnitType, TextTrimming, TextWrapping, TextAlignment,
    VerticalAlignment, HorizontalAlignment, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult,
    FontWeights, Point, Visibility
)
from System.Windows.Controls import (
    Grid, ColumnDefinition, RowDefinition, StackPanel, Orientation, TextBlock, CheckBox, Button, TextBox,
    ComboBox, RadioButton, Slider, WrapPanel, Border, DockPanel, Dock, ScrollViewer, ScrollBarVisibility, Canvas
)
from System.Windows.Controls.Primitives import Popup, PlacementMode
from System.Windows.Shapes import Line
from System.Windows.Media import (
    SolidColorBrush, Color as WpfColor, VisualBrush, VisualTreeHelper, Brushes, LinearGradientBrush, GradientStop,
    PenLineCap
)
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System.Windows.Input import Key, MouseButtonState, Cursors
from System import Uri, UriKind


# --- Console output hardening ---
# ACT's Python console stream can already be closed by the time a deferred toolbar callback
# runs, even on the very first invocation of a session: observed in practice as "SystemError:
# Cannot access a closed Stream" on the very first print statement inside _prepare_environment
# (SECTION 1), thrown from deep inside ACT's own invocation of HighFiveOut - unrelated to
# anything this script does, and not something Python code can prevent ACT from doing. Wrapping
# sys.stdout once, at the very top of HighFiveOut (SECTION 8), makes every later print in this
# script AND in the execfile'd 00 -> 05 modules degrade silently instead of crashing the whole
# report generation - consistent with this project's "log and continue" error handling (see
# CLAUDE.md).


class _SafeStdout(object):
    """Wraps a stream so that write()/flush() failures (e.g. a closed ACT console stream) are swallowed instead of raised."""

    def __init__(self, target):
        self._target = target

    def write(self, text):
        try:
            self._target.write(text)
        except Exception:
            pass

    def flush(self):
        try:
            self._target.flush()
        except Exception:
            pass


def _harden_console_output():
    """
    Does: wraps sys.stdout in _SafeStdout so a closed ACT console stream can no longer crash a print statement.
    Depends on: sys.stdout, _SafeStdout.
    Returns: nothing (side effect: reassigns sys.stdout, once).
    """
    if not isinstance(sys.stdout, _SafeStdout):
        sys.stdout = _SafeStdout(sys.stdout)


# --- SECTION 1 - Loading the project modules (00 -> 05) ---
# This whole block lives inside a function (not at module level) because ACT loads this script
# for EVERY context where the extension is registered, including the Workbench "Project" page
# before Mechanical is even open. Executable module-level code that touches ExtAPI would crash
# as soon as the extension loads (Project context) instead of waiting for the toolbar button
# click (callback HighFiveOut, SECTION 8) - this is what "Can not load extension for context
# Project" comes from. execfile() runs each file in this script's global namespace (explicit 2nd
# argument = globals()), as if its content had been copy-pasted into the console right after the
# others.


def _prepare_environment():
    """
    Does: locates PROJECT_DIR via __file__, loads modules 00 -> 05 and the default file paths.
    Depends on: __file__ (path of this script, set by ACT when it loads the <script src> from the manifest), execfile(..., globals()).
    Returns: nothing (side effect: populates the globals PROJECT_DIR, FILE_PATH_SETTINGS, _DEFAULT_FILE_PATHS and everything defined by 00_constants.py -> 05_interactive_slides.py).
    """
    global PROJECT_DIR, FILE_PATH_SETTINGS, _DEFAULT_FILE_PATHS

    # PROJECT_DIR = "Liebherr Report Generator" folder containing THIS script, wherever the
    # extension was installed (an "Additional Extension Folder" picked in Options, or the
    # extraction folder of an .actx installed via the Extension Manager, on any machine).
    #
    # Located via __file__ rather than via ExtAPI.DataModel.Project.ProjectDirectory: the latter
    # refers to the Ansys project currently OPEN in Mechanical, which has nothing to do with
    # where the extension is installed - using it would break the toolbar button as soon as it
    # is triggered in a different project than the one used at install time (the extension must
    # behave identically in any project, without copying anything into it). __file__ is set by
    # ACT when it loads the <script src> from the manifest (unlike the scripting console, where
    # __file__/os.getcwd() proved unreliable): it is the real path of this .py file as installed,
    # never the path of the project currently open.
    try:
        PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        raise RuntimeError(
            "__file__ is not defined in this execution context. This script must be loaded by "
            "ACT via <script src> in Liebherr Report Generator.xml (running it directly from the "
            "scripting console is no longer the supported installation mode)."
        )

    if not os.path.isfile(os.path.join(PROJECT_DIR, "00_constants.py")):
        raise IOError(
            "Missing modules next to AnsysReportGenerator_WPF.py: {}. The extension appears to "
            "be incompletely installed/unpacked - check that 00_constants.py -> "
            "05_interactive_slides.py, AnsysReportGenerator_WPF.xaml and the PowerPoint template "
            "are all present in this same folder.".format(PROJECT_DIR)
        )

    _MODULE_FILES = [
        "00_constants.py",
        "01_data_export.py",
        "02_image_export.py",
        "03_ppt_utils.py",
        "04_slides.py",
        "05_interactive_slides.py",
    ]

    for _module_file in _MODULE_FILES:
        _module_path = os.path.join(PROJECT_DIR, _module_file)
        if not os.path.exists(_module_path):
            raise IOError(
                "Module not found: {}. Check PROJECT_DIR at the top of "
                "AnsysReportGenerator_WPF.py.".format(_module_path)
            )
        print "Loading module: " + _module_path
        execfile(_module_path, globals())

    print "All modules loaded."

    # --- File paths editable from the "Files" tab ---
    # Original values from 00_constants.py, captured once here (before any modification from the
    # UI) so that the "Reset paths" button can always revert to them. Key = name of the
    # corresponding global in 00_constants.py, reassigned directly via globals()[key] = ...: all
    # modules 00_constants.py -> 05_interactive_slides.py read this same global at call time, no
    # other change needed elsewhere.
    FILE_PATH_SETTINGS = [
        ("TEMPLATE_PATH", "txtPathTemplate", "btnBrowseTemplate", "file"),
        ("IMAGE_EXPORT_FOLDER", "txtPathImages", "btnBrowseImages", "folder"),
        ("CSV_EXPORT_FOLDER", "txtPathCsv", "btnBrowseCsv", "folder"),
        ("LEGEND_FOLDER", "txtPathLegends", "btnBrowseLegends", "folder"),
        ("REPORT_OUTPUT_FOLDER", "txtPathReports", "btnBrowseReports", "folder"),
    ]

    _DEFAULT_FILE_PATHS = dict((name, globals()[name]) for name, _, _, _ in FILE_PATH_SETTINGS)


# --- SECTION 2 - Shared helpers (status colors, search) ---

# --- Status colors for selection rows (3 states) ---
# 3 states based on selection AND configuration:
#   - not selected (regardless of its configuration state)
#   - selected, not yet configured via "..."
#   - selected AND configured via "..."
# (see the 3 brushes below for the exact colors of each state)

ROW_STATUS_NOT_SELECTED_BRUSH = SolidColorBrush(WpfColor.FromRgb(0xFF, 0xD0, 0x00))
ROW_STATUS_SELECTED_BRUSH = SolidColorBrush(WpfColor.FromRgb(0xBE, 0xE3, 0xDB))
ROW_STATUS_CONFIGURED_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x7D, 0xCE, 0x82))


def _row_status_brush(row):
    """
    Does: determines the background color of a selection row based on its checked/configured state.
    Depends on: row.checkbox.IsChecked, row.row_config.configured, the 3 ROW_STATUS_* brushes.
    Returns: SolidColorBrush, the background color to apply to row.border.
    """
    if not row.checkbox.IsChecked:
        return ROW_STATUS_NOT_SELECTED_BRUSH
    if row.row_config.configured:
        return ROW_STATUS_CONFIGURED_BRUSH
    return ROW_STATUS_SELECTED_BRUSH


def _general_slide_status_text(row_config):
    """
    Does: builds the status text shown under the title of the Geometry/Mesh cards
    ("Overview slides" tab, tab 01) - configuration state and effective view.
    Depends on: row_config.configured/view_name.
    Returns: str, e.g. "to configure - current view" or "configured - view=ISO View".
    """
    state = "configured" if row_config.configured else "to configure"
    view = "view={}".format(row_config.view_name) if row_config.view_name else "current view"
    return "{} - {}".format(state, view)


# --- Filter by contact type (section "Contacts to display") ---
# Based on the name PREFIX (not on contact.ContactType, the Ansys API): a contact renamed by
# the engineer ("custom" name) must fall into "Autres" even if its type remains Frictional/Bonded
# on the solver side - it's the name shown in the list, not the technical type, that this filter sorts on.

CONTACTS_FILTER_OPTIONS = ["Tous", "Frictional", "Bonded", "Autres"]


def _classify_contact_name(name):
    """
    Does: classifies a Contact Region name based on its prefix ("Frictional-...", "Bonded-...", or custom).
    Depends on: nothing (string comparison, case-insensitive).
    Returns: str, "Frictional"/"Bonded"/"Autres".
    """
    lowered = (name or "").strip().lower()
    if lowered.startswith("frictional"):
        return "Frictional"
    if lowered.startswith("bonded"):
        return "Bonded"
    return "Autres"


# --- Search fields: greyed-out placeholder text ---

SEARCH_PLACEHOLDER = "Search..."
SEARCH_PLACEHOLDER_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x79, 0x7E, 0x8A))  # same grey as TextMutedBrush (xaml)
SEARCH_TEXT_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x00, 0x00, 0x00))  # same black as TextPrimaryBrush (xaml)
SEARCH_BOX_DEFAULT_BACKGROUND = SolidColorBrush(WpfColor.FromRgb(0xFA, 0xFB, 0xFC))
SEARCH_BOX_NO_MATCH_BACKGROUND = SolidColorBrush(WpfColor.FromRgb(0xF8, 0xD9, 0xDC))
SEARCH_HIGHLIGHT_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x00, 0x8D, 0xD5))

# --- Preview cards: hover color (instead of a zoom effect) ---

CARD_NORMAL_BACKGROUND = SolidColorBrush(WpfColor.FromRgb(0xFF, 0xFF, 0xFF))
CARD_HOVER_BACKGROUND = SolidColorBrush(WpfColor.FromRgb(0xE7, 0xEC, 0xF8))  # same blue as HoverBrush (xaml)

# Card width.
CARD_WIDTH = 340

# List container (see _build_preview_list_container): FIXED height (not a ceiling) so that
# all preview cards share the same size, whether they have 1 or 50 checked items -
# otherwise a card with few items (e.g. Mesh) looks tiny next to a full card (e.g.
# Boundary Conditions). Beyond this height, the list remains browsable via scrolling,
# signaled by a fade (PREVIEW_LIST_FADE_HEIGHT) at the bottom.
PREVIEW_LIST_DEFAULT_HEIGHT = 130
PREVIEW_LIST_FADE_HEIGHT = 26
PREVIEW_LIST_BACKGROUND_COLOR = WpfColor.FromRgb(0xF1, 0xF2, 0xF5)
PREVIEW_LIST_BACKGROUND = SolidColorBrush(PREVIEW_LIST_BACKGROUND_COLOR)

# Bottom fade of the checkable selection lists (tabs 01/02/03, CardBorder cards): same principle
# as PREVIEW_LIST_FADE_HEIGHT above (visible only if the list actually overflows), but
# applied via OpacityMask directly on the existing ScrollViewer rather than a separate
# overlay Border - unnecessary here since these lists always sit on a solid white CardBorder
# background (see ReportGeneratorApp._attach_list_fade).
ITEM_LIST_FADE_HEIGHT = 26

# --- Resources shared with AnsysReportGenerator_WPF.xaml ---
# The fields of the configuration panels ("..." on each row, global side panel -
# see SECTION 4/5/5bis/6 and ReportGeneratorApp._open_config_panel) are built in Python code
# (Border/TextBlock/ComboBox... created directly, not loaded from the XAML) and therefore cannot
# resolve the main window's {StaticResource ...} through the XAML markup. Rather than
# redefining these styles/colors by hand on the Python side (a source of drift with
# the .xaml observed in practice), _shared_resources references the SAME ResourceDictionary as the
# main window directly, assigned once in ReportGeneratorApp.__init__ (a single app instance
# per run, see SECTION 8) - works even once these controls later join the same
# visual tree as the main window (which is the case since the move to the side panel).

_shared_resources = None  # assigned in ReportGeneratorApp.__init__


def _make_field_label(text):
    """
    Does: creates a TextBlock used as a field label in the configuration dialogs.
    Depends on: _shared_resources["TextPrimaryBrush"].
    Returns: TextBlock, the label ready to be added to the panel.
    """
    label = TextBlock()
    label.Text = text
    label.FontWeight = FontWeights.SemiBold
    label.Foreground = _shared_resources["TextPrimaryBrush"]
    label.Margin = Thickness(0, 0, 0, 2)
    return label


def _themed_textbox():
    """
    Does: creates a TextBox styled like the generic dialog field.
    Depends on: _shared_resources["DialogTextBox"] (x:Key defined in the XAML).
    Returns: TextBox, the styled field ready to use.
    """
    box = TextBox()
    box.Style = _shared_resources["DialogTextBox"]
    return box


def _themed_button(primary=False):
    """
    Does: creates a Button styled as PrimaryButton (accent) or SecondaryButton (neutral).
    Depends on: _shared_resources["PrimaryButton"/"SecondaryButton"], same resources as the main window.
    Returns: Button, the styled button ready to use.
    """
    btn = Button()
    btn.Style = _shared_resources["PrimaryButton" if primary else "SecondaryButton"]
    return btn


def _build_close_icon(size=10, thickness=1.4):
    """
    Does: builds a small vector "x" (2 crossed lines) to use as Button.Content
    for an "x" close button - a plain TextBlock("x") is never perfectly centered
    vertically within its frame (font metrics, ascender/descender), even with
    HorizontalAlignment/VerticalAlignment=Center on the ContentPresenter.
    Depends on: Canvas/Line (System.Windows.Shapes), _shared_resources["TextPrimaryBrush"].
    Returns: Canvas, the icon ready to be assigned to Button.Content (a fresh instance on each call).
    """
    canvas = Canvas()
    canvas.Width = size
    canvas.Height = size

    for x1, y1, x2, y2 in ((0, 0, size, size), (size, 0, 0, size)):
        line = Line()
        line.X1 = x1
        line.Y1 = y1
        line.X2 = x2
        line.Y2 = y2
        line.Stroke = _shared_resources["TextPrimaryBrush"]
        line.StrokeThickness = thickness
        line.StrokeStartLineCap = PenLineCap.Round
        line.StrokeEndLineCap = PenLineCap.Round
        canvas.Children.Add(line)

    return canvas


# --- Formatted console messages (key steps) ---
# Replaces blocking dialog boxes (MessageBox) for routine events:
# a MessageBox.Show() blocks both this window AND Mechanical until it is closed
# manually, which breaks the flow when generating several reports in a row.

CONSOLE_BANNER_WIDTH = 70


def _print_console_banner(title):
    """
    Does: displays a boxed title in the Mechanical console (key generation steps).
    Depends on: CONSOLE_BANNER_WIDTH.
    Returns: nothing (side effect: prints to the console).
    """
    border = "=" * CONSOLE_BANNER_WIDTH
    print border
    print title
    print border


# --- SECTION 3 - WPF selection row (checkbox + name + config) ---

class SectionRow(object):
    """
    A WPF selection row: checkbox + object name + optional configuration
    button ("..."), linked to a row_config (SlideRowConfig / GeometryPartRowConfig /
    ContactRowConfig / MeshPartRowConfig / SolutionInfoRowConfig, see 05_interactive_slides.py).

    Attributes:
        border (Border): Outer container of the row (status background, search highlight).
        checkbox (CheckBox): "Include this slide" checkbox.
        text_block (TextBlock): Displayed name of the object.
        config_button (Button): "..." button (None if the category has nothing to configure).
        row_config: Associated configuration object.
        display_name_func (callable): Preview text function for this row_config.
        panel_kind (str): State of the global side panel to display for this row_config when
            "..." is clicked ("result"/"geometry_part"/"mesh_part"/"solution_info", None if no category).
            See ReportGeneratorApp._open_config_panel.
    """

    def __init__(self, border, checkbox, text_block, config_button, row_config,
                 display_name_func, panel_kind):
        """
        Does: stores the references to the WPF controls and the configuration associated with the row.
        Depends on: nothing (simple assignment of the received parameters).
        Returns: nothing (initializes self's attributes).
        """
        self.border = border
        self.checkbox = checkbox
        self.text_block = text_block
        self.config_button = config_button
        self.row_config = row_config
        self.display_name_func = display_name_func
        self.panel_kind = panel_kind


# --- SECTION 4 - Shared fields: view / section / scale factor / steps ---
# These fields originally lived in 4 modal dialog boxes ("..." on each selection
# row). They have been replaced with ONE global side panel (see "PARAMETERS" in the
# XAML, ReportGeneratorApp._open_config_panel and the _on_config_panel_* methods): no more
# separate windows, only 4 content "kinds" (result/geometry_part/mesh_part/solution_info)
# displayed in turn in the same panel. Each pair of functions below builds
# (_build_*) then reads back (_apply_*) a set of fields on a generic `target` (see
# _ConfigFieldsHolder): this decoupling is what allows the same code to serve both the
# global side panel AND the inline row panel of the "Combined slide" tab (_build_row_config_fields
# only, without steps - a result fixed per row).

def _build_row_config_fields(target, root, row_config, views, section_plane_labels, legend_names):
    """
    Does: builds the common graphics configuration fields (view, section, legend, appearance,
    scoping, scale factor - without the steps section) and adds them to root. Shared by the
    global side panel ("kind"="result", see ReportGeneratorApp._open_config_panel) and by the
    inline row panel of the "Combined slide" tab (a result fixed per row, never a notion of
    step - see ReportGeneratorApp._show_multi_result_editor).
    Depends on: _make_field_label, get_result_display_unit, CONTOUR_VIEW_OPTIONS/LEGEND_ORIENTATION_OPTIONS/SCOPING_DISPLAY_OPTIONS.
    Returns: nothing (sets on target: cmb_view/cmb_section/cmb_legend/cmb_contour_view/cmb_legend_orientation/cmb_scoping_display/txt_scale).
    """
    root.Children.Add(_make_field_label("View (View Manager):"))
    target.cmb_view = ComboBox()
    target.cmb_view.Margin = Thickness(0, 4, 0, 12)
    target.cmb_view.Items.Add(NO_VIEW_LABEL)
    for name in sorted(views.keys()):
        target.cmb_view.Items.Add(name)
    if row_config.view_name and row_config.view_name in views:
        target.cmb_view.SelectedItem = row_config.view_name
    else:
        target.cmb_view.SelectedIndex = 0
    root.Children.Add(target.cmb_view)

    root.Children.Add(_make_field_label("Section (Section Plane):"))
    target.cmb_section = ComboBox()
    target.cmb_section.Margin = Thickness(0, 4, 0, 12)
    target.cmb_section.Items.Add(NO_SECTION_LABEL)
    for name in section_plane_labels:
        target.cmb_section.Items.Add(name)
    if row_config.section_name and row_config.section_name in section_plane_labels:
        target.cmb_section.SelectedItem = row_config.section_name
    else:
        target.cmb_section.SelectedIndex = 0
    root.Children.Add(target.cmb_section)

    # The unit displayed here is EXACTLY the one that will be passed to ExtAPI.Graphics.ImportLegend()
    # during generation: immediate visual diagnostic, without going through the console.
    # force_evaluate=False: indicative read only (no costly re-evaluation of the
    # result) so that opening this window remains instantaneous; the actual application
    # of the legend (apply_legend_if_exists) always fully re-evaluates the result.
    detected_unit = get_result_display_unit(row_config.obj, force_evaluate=False)
    lbl_unit = TextBlock()
    lbl_unit.Text = "Unit detected for ImportLegend: " + (detected_unit if detected_unit else "none")
    lbl_unit.Foreground = _shared_resources["DiagnosticLabelBrush"]
    lbl_unit.FontWeight = FontWeights.Bold
    lbl_unit.TextWrapping = TextWrapping.Wrap
    lbl_unit.Margin = Thickness(0, 0, 0, 6)
    root.Children.Add(lbl_unit)

    root.Children.Add(_make_field_label("Legend:"))
    target.cmb_legend = ComboBox()
    target.cmb_legend.Margin = Thickness(0, 4, 0, 12)
    target.cmb_legend.Items.Add(NO_LEGEND_LABEL)
    for name in legend_names:
        target.cmb_legend.Items.Add(name)
    if row_config.legend_name and row_config.legend_name in legend_names:
        target.cmb_legend.SelectedItem = row_config.legend_name
    else:
        target.cmb_legend.SelectedIndex = 0
    root.Children.Add(target.cmb_legend)

    root.Children.Add(_make_field_label("Color display (Contour View):"))
    target.cmb_contour_view = ComboBox()
    target.cmb_contour_view.Margin = Thickness(0, 4, 0, 12)
    for label, value in CONTOUR_VIEW_OPTIONS:
        target.cmb_contour_view.Items.Add(label)
    target.cmb_contour_view.SelectedItem = contour_view_label(row_config.contour_view)
    root.Children.Add(target.cmb_contour_view)

    root.Children.Add(_make_field_label("Legend orientation:"))
    target.cmb_legend_orientation = ComboBox()
    target.cmb_legend_orientation.Margin = Thickness(0, 4, 0, 12)
    for label, value in LEGEND_ORIENTATION_OPTIONS:
        target.cmb_legend_orientation.Items.Add(label)
    target.cmb_legend_orientation.SelectedItem = legend_orientation_label(row_config.legend_orientation)
    root.Children.Add(target.cmb_legend_orientation)

    root.Children.Add(_make_field_label("Scoping display:"))
    target.cmb_scoping_display = ComboBox()
    target.cmb_scoping_display.Margin = Thickness(0, 4, 0, 12)
    for label, value in SCOPING_DISPLAY_OPTIONS:
        target.cmb_scoping_display.Items.Add(label)
    target.cmb_scoping_display.SelectedItem = scoping_display_label(row_config.scoping_display)
    root.Children.Add(target.cmb_scoping_display)

    root.Children.Add(_make_field_label("Deformation scale:"))
    target.cmb_deformation_scale_mode = ComboBox()
    target.cmb_deformation_scale_mode.Margin = Thickness(0, 4, 0, 12)
    for label, value in DEFORMATION_SCALE_MODE_OPTIONS:
        target.cmb_deformation_scale_mode.Items.Add(label)
    target.cmb_deformation_scale_mode.SelectedItem = deformation_scale_mode_label(row_config.deformation_scale_mode)
    root.Children.Add(target.cmb_deformation_scale_mode)

    root.Children.Add(_make_field_label("Deformation scale factor (Manual mode only, default = 1):"))
    target.txt_scale = _themed_textbox()
    target.txt_scale.Width = 100
    target.txt_scale.HorizontalAlignment = HorizontalAlignment.Left
    target.txt_scale.Margin = Thickness(0, 4, 0, 12)
    target.txt_scale.Text = "1" if row_config.scale_factor == 1.0 else str(row_config.scale_factor)
    root.Children.Add(target.txt_scale)


def _apply_row_config_fields(target, row_config):
    """
    Does: reads the common fields (view/section/legend/appearance/scoping/scale factor) from target
    and applies them to row_config. Shared by the global side panel (_on_config_panel_apply) and
    by the inline row panel of the "Combined slide" tab (see _build_row_config_fields).
    Depends on: target.cmb_view/cmb_section/cmb_legend/cmb_contour_view/cmb_legend_orientation/cmb_scoping_display/
        cmb_deformation_scale_mode/txt_scale.
    Returns: nothing (side effect on row_config only; does not touch row_config.configured, nor the steps).
    """
    selected_view = unicode(target.cmb_view.SelectedItem)
    row_config.view_name = None if selected_view == NO_VIEW_LABEL else selected_view

    selected_section = unicode(target.cmb_section.SelectedItem)
    row_config.section_name = None if selected_section == NO_SECTION_LABEL else selected_section

    selected_legend = unicode(target.cmb_legend.SelectedItem)
    row_config.legend_name = None if selected_legend == NO_LEGEND_LABEL else selected_legend

    row_config.contour_view = contour_view_from_label(unicode(target.cmb_contour_view.SelectedItem))
    row_config.legend_orientation = legend_orientation_from_label(unicode(target.cmb_legend_orientation.SelectedItem))
    row_config.scoping_display = scoping_display_from_label(unicode(target.cmb_scoping_display.SelectedItem))
    row_config.deformation_scale_mode = deformation_scale_mode_from_label(
        unicode(target.cmb_deformation_scale_mode.SelectedItem))

    # The field's value is only read/validated in Manual mode: in Auto Scale x1/x2 mode, the
    # applied multiplier is fixed (see apply_scale_factor), this field is ignored - no point
    # warning about an invalid value that won't be used anyway.
    if row_config.deformation_scale_mode == "manual":
        try:
            scale_value = float(target.txt_scale.Text.strip().replace(",", "."))
            if scale_value <= 0:
                raise ValueError("Scale factor must be positive.")
            row_config.scale_factor = scale_value
        except ValueError:
            row_config.scale_factor = 1.0
            MessageBox.Show("Invalid scale factor value: the default value (1) has been applied.",
                             "Invalid scale factor", MessageBoxButton.OK, MessageBoxImage.Warning)


def _build_steps_section_fields(target, root, row_config, step_count):
    """
    Does: builds the "Loadcases" section (steps + individual/combined display mode) and adds it to root.
    Depends on: row_config.selected_steps/step_display_mode, _shared_resources.
    Returns: nothing (sets on target: step_checkboxes/radio_individual/radio_combined).
    """
    group = Border()
    group.BorderBrush = _shared_resources["CardBorderBrush"]
    group.BorderThickness = Thickness(1)
    group.CornerRadius = CornerRadius(0)
    group.Padding = Thickness(10)
    group.Margin = Thickness(0, 4, 0, 0)

    panel = StackPanel()
    group.Child = panel

    lbl_info = TextBlock()
    lbl_info.Text = "Available loadcases: {}".format(step_count)
    lbl_info.FontWeight = FontWeights.SemiBold
    lbl_info.Margin = Thickness(0, 0, 0, 6)
    panel.Children.Add(lbl_info)

    wrap = WrapPanel()
    selected_steps = row_config.selected_steps or []
    target.step_checkboxes = []
    for step in range(1, step_count + 1):
        cb = CheckBox()
        cb.Content = "Step {}".format(step)
        cb.Tag = step
        cb.IsChecked = step in selected_steps
        cb.Width = 90
        cb.Margin = Thickness(0, 2, 10, 2)
        wrap.Children.Add(cb)
        target.step_checkboxes.append(cb)
    panel.Children.Add(wrap)

    lbl_hint = TextBlock()
    lbl_hint.Text = "(none checked = current state)"
    lbl_hint.FontSize = 11
    lbl_hint.Foreground = SEARCH_PLACEHOLDER_BRUSH
    lbl_hint.Margin = Thickness(0, 6, 0, 2)
    panel.Children.Add(lbl_hint)

    steps_buttons = StackPanel()
    steps_buttons.Orientation = Orientation.Horizontal
    steps_buttons.Margin = Thickness(0, 0, 0, 10)

    def _on_select_all_steps(sender, e):
        for cb in target.step_checkboxes:
            cb.IsChecked = True

    def _on_deselect_all_steps(sender, e):
        for cb in target.step_checkboxes:
            cb.IsChecked = False

    btn_select_all_steps = _themed_button()
    btn_select_all_steps.Content = "Select all"
    btn_select_all_steps.Padding = Thickness(8, 2, 8, 2)
    btn_select_all_steps.FontSize = 11
    btn_select_all_steps.Click += _on_select_all_steps
    steps_buttons.Children.Add(btn_select_all_steps)

    btn_deselect_all_steps = _themed_button()
    btn_deselect_all_steps.Content = "Deselect all"
    btn_deselect_all_steps.Padding = Thickness(8, 2, 8, 2)
    btn_deselect_all_steps.FontSize = 11
    btn_deselect_all_steps.Margin = Thickness(6, 0, 0, 0)
    btn_deselect_all_steps.Click += _on_deselect_all_steps
    steps_buttons.Children.Add(btn_deselect_all_steps)

    panel.Children.Add(steps_buttons)

    target.radio_individual = RadioButton()
    target.radio_individual.Content = "Individual slides (1 per step)"
    target.radio_individual.GroupName = "StepDisplayMode"
    target.radio_individual.IsChecked = (row_config.step_display_mode != "combined")
    target.radio_individual.Margin = Thickness(0, 0, 0, 2)
    panel.Children.Add(target.radio_individual)

    target.radio_combined = RadioButton()
    target.radio_combined.Content = "Combined slide (if template available)"
    target.radio_combined.GroupName = "StepDisplayMode"
    target.radio_combined.IsChecked = (row_config.step_display_mode == "combined")
    panel.Children.Add(target.radio_combined)

    root.Children.Add(group)


def _apply_steps_section_fields(target, row_config):
    """
    Does: reads the step selection and display mode from target and applies them to row_config.
    Depends on: target.step_checkboxes/radio_combined, get_multi_step_template, MULTI_STEP_SLIDE_TEMPLATES.
    Returns: nothing (side effect on row_config.selected_steps/step_display_mode).
    """
    checked_steps = [cb.Tag for cb in target.step_checkboxes if cb.IsChecked]
    row_config.selected_steps = checked_steps if checked_steps else None

    if not checked_steps:
        row_config.step_display_mode = "individual"
    elif target.radio_combined.IsChecked:
        if get_multi_step_template(len(checked_steps)):
            row_config.step_display_mode = "combined"
        else:
            row_config.step_display_mode = "individual"
            supported = ", ".join(str(n) for n in sorted(MULTI_STEP_SLIDE_TEMPLATES.keys()))
            MessageBox.Show(
                "No combined slide template exists for {} step(s) (supported counts: {}). "
                "Individual slides will be generated instead.".format(len(checked_steps), supported),
                "Combined mode unavailable", MessageBoxButton.OK, MessageBoxImage.Warning
            )
    else:
        row_config.step_display_mode = "individual"


# --- SECTION 5 - Shared fields: view / section / opacity (geometry per part) ---

def _build_geometry_part_fields(target, root, row_config, views, section_plane_labels):
    """
    Does: builds the view/section/context opacity fields for an isolated part (geometry) and adds them to root.
    Depends on: _make_field_label.
    Returns: nothing (sets on target: cmb_view/cmb_section/slider_opacity/lbl_opacity_value).
    """
    root.Children.Add(_make_field_label("View (View Manager):"))
    target.cmb_view = ComboBox()
    target.cmb_view.Margin = Thickness(0, 4, 0, 12)
    target.cmb_view.Items.Add(NO_VIEW_LABEL)
    for name in sorted(views.keys()):
        target.cmb_view.Items.Add(name)
    if row_config.view_name and row_config.view_name in views:
        target.cmb_view.SelectedItem = row_config.view_name
    else:
        target.cmb_view.SelectedIndex = 0
    root.Children.Add(target.cmb_view)

    root.Children.Add(_make_field_label("Section (Section Plane):"))
    target.cmb_section = ComboBox()
    target.cmb_section.Margin = Thickness(0, 4, 0, 12)
    target.cmb_section.Items.Add(NO_SECTION_LABEL)
    for name in section_plane_labels:
        target.cmb_section.Items.Add(name)
    if row_config.section_name and row_config.section_name in section_plane_labels:
        target.cmb_section.SelectedItem = row_config.section_name
    else:
        target.cmb_section.SelectedIndex = 0
    root.Children.Add(target.cmb_section)

    root.Children.Add(_make_field_label("Context opacity (other parts):"))

    opacity_row = StackPanel()
    opacity_row.Orientation = Orientation.Horizontal
    opacity_row.Margin = Thickness(0, 4, 0, 12)

    target.slider_opacity = Slider()
    target.slider_opacity.Minimum = 0
    target.slider_opacity.Maximum = 100
    target.slider_opacity.TickFrequency = 10
    target.slider_opacity.Width = 260
    target.slider_opacity.Value = row_config.context_opacity_percent
    opacity_row.Children.Add(target.slider_opacity)

    target.lbl_opacity_value = TextBlock()
    target.lbl_opacity_value.Text = "{} %".format(row_config.context_opacity_percent)
    target.lbl_opacity_value.Margin = Thickness(10, 0, 0, 0)
    target.lbl_opacity_value.VerticalAlignment = VerticalAlignment.Center
    opacity_row.Children.Add(target.lbl_opacity_value)

    def _on_opacity_changed(sender, e):
        target.lbl_opacity_value.Text = "{} %".format(int(target.slider_opacity.Value))
    target.slider_opacity.ValueChanged += _on_opacity_changed

    root.Children.Add(opacity_row)


def _apply_geometry_part_fields(target, row_config):
    """
    Does: reads the view/section/opacity fields from target and applies them to row_config.
    Depends on: target.cmb_view/cmb_section/slider_opacity.
    Returns: nothing (side effect on row_config).
    """
    selected_view = unicode(target.cmb_view.SelectedItem)
    row_config.view_name = None if selected_view == NO_VIEW_LABEL else selected_view

    selected_section = unicode(target.cmb_section.SelectedItem)
    row_config.section_name = None if selected_section == NO_SECTION_LABEL else selected_section

    row_config.context_opacity_percent = int(target.slider_opacity.Value)


# --- SECTION 5bis - Shared fields: view (isolated part geometry, mesh) ---

def _build_mesh_part_fields(target, root, row_config, views):
    """
    Does: builds the view field (View Manager) only and adds it to root - no section or
    opacity (unlike _build_geometry_part_fields): isolation is done by fully hiding
    the other bodies (show_only_body), a context section/opacity would not make sense
    here. Reused as-is for Geometry/Mesh/Analysis context (view only there too).
    Depends on: _make_field_label.
    Returns: nothing (sets on target: cmb_view).
    """
    root.Children.Add(_make_field_label("View (View Manager):"))
    target.cmb_view = ComboBox()
    target.cmb_view.Margin = Thickness(0, 4, 0, 12)
    target.cmb_view.Items.Add(NO_VIEW_LABEL)
    for name in sorted(views.keys()):
        target.cmb_view.Items.Add(name)
    if row_config.view_name and row_config.view_name in views:
        target.cmb_view.SelectedItem = row_config.view_name
    else:
        target.cmb_view.SelectedIndex = 0
    root.Children.Add(target.cmb_view)


def _apply_mesh_part_fields(target, row_config):
    """
    Does: reads the view field from target and applies it to row_config.
    Depends on: target.cmb_view.
    Returns: nothing (side effect on row_config).
    """
    selected_view = unicode(target.cmb_view.SelectedItem)
    row_config.view_name = None if selected_view == NO_VIEW_LABEL else selected_view


# --- SECTION 6 - Shared fields: title / axes / color (Solution Information) ---

def _build_solution_info_fields(target, root, row_config):
    """
    Does: builds the title/axes/curve color fields for a Solution Information tracker and adds them to root.
    Depends on: _make_field_label, _themed_textbox, CURVE_COLOR_OPTIONS, curve_color_label.
    Returns: nothing (sets on target: txt_title/txt_x_label/txt_y_label/cmb_color).
    """
    root.Children.Add(_make_field_label("Chart title (empty = tracker name):"))
    target.txt_title = _themed_textbox()
    target.txt_title.Margin = Thickness(0, 4, 0, 12)
    target.txt_title.Text = row_config.chart_title or ""
    root.Children.Add(target.txt_title)

    root.Children.Add(_make_field_label("X axis name (empty = inferred from CSV):"))
    target.txt_x_label = _themed_textbox()
    target.txt_x_label.Margin = Thickness(0, 4, 0, 12)
    target.txt_x_label.Text = row_config.x_axis_label or ""
    root.Children.Add(target.txt_x_label)

    root.Children.Add(_make_field_label("Y axis name (empty = inferred from CSV, single curve):"))
    target.txt_y_label = _themed_textbox()
    target.txt_y_label.Margin = Thickness(0, 4, 0, 12)
    target.txt_y_label.Text = row_config.y_axis_label or ""
    root.Children.Add(target.txt_y_label)

    root.Children.Add(_make_field_label("Curve color:"))
    target.cmb_color = ComboBox()
    target.cmb_color.Margin = Thickness(0, 4, 0, 12)
    for color_label, color_value in CURVE_COLOR_OPTIONS:
        target.cmb_color.Items.Add(color_label)
    target.cmb_color.SelectedItem = curve_color_label(row_config.curve_color)
    root.Children.Add(target.cmb_color)


def _apply_solution_info_fields(target, row_config):
    """
    Does: reads the title/axes/color fields from target and applies them to row_config.
    Depends on: target.txt_title/txt_x_label/txt_y_label/cmb_color, curve_color_from_label.
    Returns: nothing (side effect on row_config).
    """
    row_config.chart_title = target.txt_title.Text.strip() or None
    row_config.x_axis_label = target.txt_x_label.Text.strip() or None
    row_config.y_axis_label = target.txt_y_label.Text.strip() or None
    row_config.curve_color = curve_color_from_label(unicode(target.cmb_color.SelectedItem))


# --- SECTION 6bis - "Combined slide (different results)": state and constants ---
# The "template then grid then cell" flow used to live in 3 modal dialog boxes
# (MultiResultTemplatePickerWindow / MultiResultGridWindow / ResultPickerWindow). It is now
# integrated directly into the "04   Combined slide" tab of the main window (see
# ReportGeneratorApp._build_multi_result_tab and the _multi_result_*/_on_multi_result_* methods):
# no more separate window, the template choice, the grid, and a cell's configuration all live
# in this same tab (template at the top, grid on the left, cell panel on the right).

GRID_CELL_UNCONFIGURED_BRUSH = ROW_STATUS_NOT_SELECTED_BRUSH
GRID_CELL_CONFIGURED_BRUSH = ROW_STATUS_CONFIGURED_BRUSH
GRID_CELL_DISABLED_BRUSH = SolidColorBrush(WpfColor.FromRgb(0xD8, 0xD8, 0xD8))
GRID_CELL_SELECTED_BORDER_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x00, 0x8D, 0xD5))

MULTI_RESULT_CELL_TOTAL = 8  # number of grid slots (2 rows x 4 columns, see gridMultiResultCells)


class MultiResultSlideConfig(object):
    """
    A "different results" combined slide configured from the "Combined slide" tab, awaiting
    generation: it is not built immediately but stored in
    ReportGeneratorApp._multi_result_slides and appears as a full-fledged card in the
    Preview tab, generated only when clicking "Generate report" (same PowerPoint session as
    everything else, same drag-and-drop order).
    """

    def __init__(self, template_count, cell_configs):
        """
        Does: stores the number of slots and the graphics configuration of each cell.
        Depends on: nothing (simple assignments).
        Returns: nothing (constructor).
        """
        self.template_count = template_count
        self.cell_configs = cell_configs  # list of SlideRowConfig, one per cell, in slot order


class _ConfigFieldsHolder(object):
    """
    Generic container for the controls of a configuration panel built in code (Border/
    ComboBox/TextBox... without a dedicated window): serves as the `target` for _build_row_config_fields /
    _build_steps_section_fields / _build_geometry_part_fields / _build_mesh_part_fields /
    _build_solution_info_fields (and their corresponding _apply_*), both for the
    global side panel ("..." on any row, see ReportGeneratorApp._open_config_panel)
    and for the cell panel of the "Combined slide" tab (_show_multi_result_editor).
    """
    pass


# --- SECTION 7 - Main window (loaded from AnsysReportGenerator_WPF.xaml) ---

class ReportGeneratorApp(object):
    """
    WPF entry point of the application: same business logic (functions from 04_slides.py /
    05_interactive_slides.py), presentation loaded from AnsysReportGenerator_WPF.xaml and
    organized into 3 tabs:

    - "General slides": Geometry / Mesh (simple checkboxes) +
      "Parts to isolate (geometry)", "Part to isolate mesh" and "Analysis
      context" (one "Analysis Parameters" slide per checked analysis - see
      collect_analyses in 05_interactive_slides.py) (selection grids).
    - "Conditions and contacts": Boundary Conditions, Bolt Pretension,
      Contacts to display, Connection: Contact Tool (Contact Tool without step,
      Connections branch), Solution Information.
    - "Result categories": Contact Tool Results (Contact Tool with
      steps, Solution branch), Results, Bolt Tool - one slide per checked
      row.

    A 4th tab "Report preview" displays one card per selected slide
    (name + detailed parameters), reorderable via drag-and-drop: the chosen
    order is respected at generation time (see self._preview_order).
    """

    def __init__(self, xaml_path):
        """
        Does: initializes the application (collecting Mechanical data, loading the XAML, wiring the controls).
        Depends on: collect_* (05_interactive_slides.py), ExtAPI.DataModel, self._load_window/_find_controls/_build_sections/_wire_*.
        Returns: nothing (builds self.window ready to be shown by SECTION 8).
        """
        remove_stale_figures()

        self._analysis = ExtAPI.DataModel.Project.Model.Analyses[0]

        # Results / Contact Tool Results / Bolt Tool / Bolt Pretension / Solution Information
        # are compiled from ALL analyses in the project as soon as there is more than one, and
        # tagged with their originating analysis (tuples (obj, analysis)) to differentiate
        # identical names coming from two different analyses (see analysis_suffix). On a
        # single-analysis project, behavior identical to before (lists tagged (obj, None), never suffixed).
        self._analyses = collect_analyses()
        self._multi_analysis = len(self._analyses) > 1
        print "Analyses found: {}".format(len(self._analyses))
        for _i, _a in enumerate(self._analyses, start=1):
            print "  {} : {}".format(_i, _a.Name)

        self._contact_tool_connections_results = collect_connection_contact_tool_results(self._analysis)
        self._bodies = collect_bodies()
        self._contact_regions = collect_contact_regions()

        if self._multi_analysis:
            self._bcs = collect_boundary_conditions_multi(self._analyses)
            self._results = collect_all_results_multi(self._analyses)
            self._contact_tool_results = collect_contact_tool_results_multi(self._analyses)
            self._bolt_tool_results = collect_bolt_tool_results_multi(self._analyses)
            self._bolt_pretensions = collect_bolt_pretensions_multi(self._analyses)
            self._solution_info_trackers = collect_solution_information_trackers_multi(self._analyses)
        else:
            self._bcs = [(obj, None) for obj in collect_boundary_conditions(self._analysis)]
            self._results = [(obj, None) for obj in collect_all_results(self._analysis)]
            self._contact_tool_results = [(obj, None) for obj in collect_contact_tool_results(self._analysis)]
            self._bolt_tool_results = [(obj, None) for obj in collect_bolt_tool_results(self._analysis)]
            self._bolt_pretensions = [(obj, None) for obj in collect_bolt_pretensions(self._analysis)]
            self._solution_info_trackers = [(obj, None) for obj in collect_solution_information_trackers(self._analysis)]

        self._views = collect_views()
        self._section_planes = collect_section_planes()
        self._section_plane_labels = [
            section_plane_label(sp, i) for i, sp in enumerate(self._section_planes)
        ]
        self._step_count = get_step_count(self._analysis)
        self._legend_names = collect_legend_files()

        # View (View Manager) chosen for the Geometry/Mesh slides (simple checkboxes,
        # no list): reuses MeshPartRowConfig (view only, already used for isolated-part
        # mesh) rather than introducing a dedicated class, its attributes (obj/view_name/
        # configured) being sufficient as-is.
        self._geometry_view_config = MeshPartRowConfig(ExtAPI.DataModel.Project.Model.Geometry)
        self._mesh_view_config = MeshPartRowConfig(ExtAPI.DataModel.Project.Model.Mesh)

        # Choice of mesh table (default / full) for the Mesh slide: moved from the
        # ComboBox formerly in the "Mesh slide" card (tab 01) to the "PARAMETERS" panel
        # of that same row (see _open_config_panel, kind="mesh_part" + row_config is self._mesh_view_config).
        self._mesh_table_full = False  # False = default table, True = full table

        self._sections = {}       # section name -> {rows, group_key, label, search_box}
        self._section_order = []  # stable display order for the preview
        self._multi_result_slides = []  # list of MultiResultSlideConfig, one per "different results" slide added to the report (see "Combined slide" tab, _on_multi_result_add_to_report)

        # --- State of the global configuration side panel ("...", see _open_config_panel) ---
        self._config_panel_kind = None         # "result"/"geometry_part"/"mesh_part"/"solution_info"/None
        self._config_panel_row_config = None   # row_config currently being edited in the panel
        self._config_panel_fields = None       # current _ConfigFieldsHolder (cmb_view/cmb_section/...)
        self._config_panel_refresh = None      # callable() invoked after "Apply" (refreshes the row/preview)
        self._config_panel_bulk_rows = None    # list of SectionRow if panel opened in bulk mode, None otherwise

        # --- State of the "Combined slide" tab (grid under construction, see _build_multi_result_tab) ---
        self._mr_template_count = None    # number of active cells in the chosen template (2/3/4/6/8)
        self._mr_cell_configs = [None] * MULTI_RESULT_CELL_TOTAL  # index -> SlideRowConfig or None
        self._mr_cell_borders = []        # Border of each cell of gridMultiResultCells, in order
        self._mr_cell_labels = []         # TextBlock of each cell, in order
        self._mr_template_buttons = {}    # result count -> Button (for the "selected" visual state)
        self._mr_selected_cell_index = None  # cell currently shown in panelMultiResultSidePanel (None = none)
        self._mr_editing = None           # (index, cfg) currently being edited in the right-hand panel ("editor" state)
        self._mr_editor_fields = None     # current _ConfigFieldsHolder (cmb_view/cmb_section/... of the displayed editor)
        self._mr_picker_rows = []         # [(Border, TextBlock, result)] of the displayed "choose a result" list

        self._preview_order = []  # ordered list of (kind, payload) tuples, reorderable via drag-and-drop
        self._entry_to_card = {}  # (kind, payload) -> Border currently displayed in panelPreview
        self._entry_to_badge = {}  # (kind, payload) -> TextBlock of the card's numbered badge

        # State of the ongoing drag-and-drop (see _begin_potential_drag / _start_drag / _end_drag).
        self._drag_pending_card = None
        self._drag_pending_entry = None
        self._drag_start_point = None
        self._drag_active = False
        self._drag_entry = None
        self._drag_source_card = None
        self._drag_popup = None

        self._last_report_path = None  # path of the last generated PPTX report (Files tab)

        self.window = self._load_window(xaml_path)

        # Makes the main window's resources (brushes + styles) accessible to the "..."
        # configuration panels built in code (see _shared_resources, section 2): a single
        # ReportGeneratorApp instance per run (SECTION 8), so this one-time assignment
        # is sufficient for the whole duration of the script.
        global _shared_resources
        _shared_resources = self.window.Resources

        self._find_controls()
        self._refresh_general_slide_status()
        self._build_sections()
        self._build_multi_result_tab()
        self._wire_contacts_filter()
        self._wire_zone_select_buttons()
        self._wire_file_paths()
        self._refresh_csv_files()
        self._refresh_data_cleanup_tiles()
        self._wire_events()
        self._update_preview()

    # --- Loading the XAML ---

    def _load_window(self, xaml_path):
        """
        Does: loads AnsysReportGenerator_WPF.xaml and builds the corresponding WPF Window.
        Depends on: StreamReader, XmlReader, XamlReader.Load (System.Windows.Markup).
        Returns: Window, the loaded window (not yet shown).
        """
        reader = StreamReader(xaml_path)
        xml_reader = XmlReader.Create(reader)
        return XamlReader.Load(xml_reader)

    def _find_controls(self):
        """
        Does: retrieves the references to the named controls (x:Name) defined in the XAML.
        Depends on: self.window.FindName, the x:Name declared in AnsysReportGenerator_WPF.xaml.
        Returns: nothing (initializes the self.btn_.../chk_.../panel_... attributes etc.).
        """
        w = self.window
        self._load_logo()

        self.btn_delete_figures = w.FindName("btnDeleteFigures")
        self.btn_reset_legends = w.FindName("btnResetLegends")
        self.btn_create_views = w.FindName("btnCreateViews")
        self.btn_export_3d = w.FindName("btnExport3D")

        self.border_config_panel = w.FindName("borderConfigPanel")
        self.panel_config_panel = w.FindName("panelConfigPanel")

        self.panel_multi_result_template_buttons = w.FindName("panelMultiResultTemplateButtons")
        self.lbl_multi_result_fill_count = w.FindName("lblMultiResultFillCount")
        self.btn_multi_result_add_to_report = w.FindName("btnMultiResultAddToReport")
        self.grid_multi_result_cells = w.FindName("gridMultiResultCells")
        self.lbl_multi_result_hint = w.FindName("lblMultiResultHint")
        self.panel_multi_result_side = w.FindName("panelMultiResultSidePanel")

        self.chk_geometry = w.FindName("chkGeometry")
        self.btn_geometry_view = w.FindName("btnGeometryView")
        self.lbl_geometry_status = w.FindName("lblGeometryStatus")
        self.chk_mesh = w.FindName("chkMesh")
        self.btn_mesh_view = w.FindName("btnMeshView")
        self.lbl_mesh_status = w.FindName("lblMeshStatus")
        self.cmb_contacts_filter = w.FindName("cmbContactsFilter")

        self.btn_check_all_general = w.FindName("btnCheckAllGeneral")
        self.btn_uncheck_all_general = w.FindName("btnUncheckAllGeneral")
        self.btn_check_all_conditions = w.FindName("btnCheckAllConditions")
        self.btn_uncheck_all_conditions = w.FindName("btnUncheckAllConditions")
        self.btn_check_all_results = w.FindName("btnCheckAllResults")
        self.btn_uncheck_all_results = w.FindName("btnUncheckAllResults")

        self.panel_preview = w.FindName("panelPreview")
        self.btn_generate = w.FindName("btnGenerate")
        self.btn_close = w.FindName("btnClose")

        self.panel_csv_files = w.FindName("panelCsvFiles")
        self.panel_data_cleanup = w.FindName("panelDataCleanup")
        self.btn_delete_all_data = w.FindName("btnDeleteAllData")
        self.btn_reset_file_paths = w.FindName("btnResetFilePaths")
        self.border_report_status = w.FindName("borderReportStatus")
        self.border_report_tile = w.FindName("borderReportTile")
        self.lbl_progress_status = w.FindName("lblProgressStatus")
        self.progress_track = w.FindName("progressTrack")
        self.progress_fill = w.FindName("progressFill")
        self.lbl_report_name = w.FindName("lblReportName")
        self.btn_report_view = w.FindName("btnReportView")
        self.btn_report_show_in_folder = w.FindName("btnReportShowInFolder")

    def _load_logo(self):
        """
        Does: loads the company logo (LOGO_PATH) into the "SidebarLogoBitmap" resource used
        by the credit card at the bottom of the tab column (see the TabControl ControlTemplate
        in the XAML, Image with Source="{DynamicResource SidebarLogoBitmap}") - DynamicResource
        (not x:Name/FindName) because this Image lives in the ControlTemplate's private NameScope,
        inaccessible from self.window.FindName().
        Depends on: self.window.Resources, LOGO_PATH (00_constants.py), BitmapImage/BitmapCacheOption/Uri.
        Returns: nothing (side effect on self.window.Resources["SidebarLogoBitmap"]; does nothing if LOGO_PATH is missing).
        """
        if not os.path.isfile(LOGO_PATH):
            return
        try:
            bitmap = BitmapImage()
            bitmap.BeginInit()
            # CacheOption.OnLoad: loads the file entirely into memory then releases the handle
            # immediately (otherwise the PNG stays open/locked by the Mechanical process for
            # the whole session).
            bitmap.CacheOption = BitmapCacheOption.OnLoad
            bitmap.UriSource = Uri(LOGO_PATH, UriKind.Absolute)
            bitmap.EndInit()
            self.window.Resources["SidebarLogoBitmap"] = bitmap
        except Exception as e:
            print "Unable to load the logo: " + str(e)

    def _refresh_general_slide_status(self):
        """
        Does: refreshes the status text of the "Geometry slide"/"Mesh slide" cards (tab 01).
        Depends on: self.lbl_geometry_status/lbl_mesh_status, self._geometry_view_config/_mesh_view_config, _general_slide_status_text.
        Returns: nothing (side effect on lblGeometryStatus/lblMeshStatus).
        """
        self.lbl_geometry_status.Text = _general_slide_status_text(self._geometry_view_config)
        self.lbl_mesh_status.Text = _general_slide_status_text(self._mesh_view_config)

    def _wire_contacts_filter(self):
        """
        Does: populates and wires the contact-type filter ComboBox (section "Contacts to display").
        Depends on: self.cmb_contacts_filter, CONTACTS_FILTER_OPTIONS, self._on_contacts_filter_changed.
        Returns: nothing (side effect on self.cmb_contacts_filter).
        """
        for label in CONTACTS_FILTER_OPTIONS:
            self.cmb_contacts_filter.Items.Add(label)
        self.cmb_contacts_filter.SelectedIndex = 0
        self.cmb_contacts_filter.SelectionChanged += self._on_contacts_filter_changed

    def _on_contacts_filter_changed(self, sender, e):
        """
        Does: VISUALLY reorders panelContacts to bring contacts of the chosen type
        (Frictional/Bonded/Autres) to the top, or restores the original order ("Tous"). Same principle as
        _perform_search (see further below): only the DISPLAY order changes, self._sections["Contacts"]["rows"]
        keeps its original order, which remains the one used for report generation.
        Depends on: self._sections["Contacts"]["rows"/"panel"], _classify_contact_name, CONTACTS_FILTER_OPTIONS.
        Returns: nothing (side effect on the visual order of panelContacts).
        """
        selected = unicode(self.cmb_contacts_filter.SelectedItem) if self.cmb_contacts_filter.SelectedItem else "Tous"
        section = self._sections["Contacts"]
        rows = section["rows"]
        panel = section["panel"]

        if selected not in CONTACTS_FILTER_OPTIONS or selected == "Tous":
            ordered_rows = rows
        else:
            matching = [row for row in rows if _classify_contact_name(row.row_config.obj.Name) == selected]
            rest = [row for row in rows if row not in matching]
            ordered_rows = matching + rest

        panel.Children.Clear()
        for row in ordered_rows:
            panel.Children.Add(row.border)

    # --- "Files" tab: editable paths ---
    # See FILE_PATH_SETTINGS / _DEFAULT_FILE_PATHS (SECTION 1): each row directly modifies
    # the corresponding global in 00_constants.py, read by all Report
    # Generator/*.py at call time - no other change needed elsewhere.

    def _wire_file_paths(self):
        """
        Does: initializes the 5 path TextBoxes with their current value and wires the "..."/reset buttons.
        Depends on: FILE_PATH_SETTINGS, self._make_path_edit_handler/_make_path_browse_handler/_on_reset_file_paths.
        Returns: nothing (side effect on the Files tab controls).
        """
        self._path_textboxes = {}
        for name, textbox_id, browse_id, kind in FILE_PATH_SETTINGS:
            textbox = self.window.FindName(textbox_id)
            browse_button = self.window.FindName(browse_id)
            textbox.Text = globals()[name]
            self._path_textboxes[name] = textbox
            textbox.LostFocus += self._make_path_edit_handler(name, kind)
            browse_button.Click += self._make_path_browse_handler(name, kind)
        self.btn_reset_file_paths.Click += self._on_reset_file_paths
        self.btn_delete_all_data.Click += self._on_delete_all_data

    def _make_path_edit_handler(self, name, kind):
        """
        Does: closes over name/kind by value to produce the LostFocus handler of a path TextBox.
        Depends on: self._apply_file_path_edit.
        Returns: function, the handler(sender, e) to wire on textbox.LostFocus.
        """
        def handler(sender, e):
            """
            Does: validates the typed path as soon as the TextBox loses focus.
            Depends on: self._apply_file_path_edit, name/kind captured by the closure.
            Returns: nothing (side effect on the corresponding global).
            """
            self._apply_file_path_edit(name, kind)
        return handler

    def _make_path_browse_handler(self, name, kind):
        """
        Does: closes over name/kind by value to produce the Click handler of the "..." button of a path row.
        Depends on: self._browse_file_path.
        Returns: function, the handler(sender, e) to wire on browse_button.Click.
        """
        def handler(sender, e):
            """
            Does: opens the browse dialog to choose a custom path.
            Depends on: self._browse_file_path, name/kind captured by the closure.
            Returns: nothing (side effect on the corresponding global).
            """
            self._browse_file_path(name, kind)
        return handler

    def _apply_file_path_edit(self, name, kind):
        """
        Does: validates the text typed in the 'name' path TextBox and reassigns the corresponding global if valid.
        Depends on: self._path_textboxes, os.path, ensure_folder_exists, globals().
        Returns: nothing (side effect: updates globals()[name] or restores the previous text).
        """
        textbox = self._path_textboxes[name]
        new_value = textbox.Text.strip()
        current_value = globals()[name]
        if new_value == current_value:
            return

        if kind == "file":
            if not os.path.isfile(new_value) or not new_value.lower().endswith(".pptx"):
                MessageBox.Show("Invalid path: the template must be an existing .pptx file.",
                                 "Invalid path", MessageBoxButton.OK, MessageBoxImage.Warning)
                textbox.Text = current_value
                return
        else:
            try:
                ensure_folder_exists(new_value)
            except Exception as ex:
                MessageBox.Show("Unable to use this folder:\n" + str(ex),
                                 "Invalid path", MessageBoxButton.OK, MessageBoxImage.Warning)
                textbox.Text = current_value
                return

        globals()[name] = new_value
        print "Path '{}' updated: {}".format(name, new_value)
        if name == "CSV_EXPORT_FOLDER":
            self._refresh_csv_files()

    def _browse_file_path(self, name, kind):
        """
        Does: opens an OpenFileDialog (template) or FolderBrowserDialog (folders) to choose a path.
        Depends on: self._path_textboxes, System.Windows.Forms (SWF.OpenFileDialog/FolderBrowserDialog), self._apply_file_path_edit.
        Returns: nothing (side effect on the TextBox and the corresponding global if a path is chosen).
        """
        textbox = self._path_textboxes[name]
        current_value = globals()[name]
        if kind == "file":
            dialog = SWF.OpenFileDialog()
            dialog.Filter = "PowerPoint (*.pptx)|*.pptx"
            if os.path.isfile(current_value):
                dialog.InitialDirectory = os.path.dirname(current_value)
            if dialog.ShowDialog() == SWF.DialogResult.OK:
                textbox.Text = dialog.FileName
                self._apply_file_path_edit(name, kind)
        else:
            dialog = SWF.FolderBrowserDialog()
            if os.path.isdir(current_value):
                dialog.SelectedPath = current_value
            if dialog.ShowDialog() == SWF.DialogResult.OK:
                textbox.Text = dialog.SelectedPath
                self._apply_file_path_edit(name, kind)

    def _on_reset_file_paths(self, sender, e):
        """
        Does: "Reset paths" button - reverts to the original values from 00_constants.py.
        Depends on: FILE_PATH_SETTINGS, _DEFAULT_FILE_PATHS, self._refresh_csv_files.
        Returns: nothing (side effect: reassigns globals() and the path TextBoxes).
        """
        for name, textbox_id, browse_id, kind in FILE_PATH_SETTINGS:
            default_value = _DEFAULT_FILE_PATHS[name]
            globals()[name] = default_value
            self._path_textboxes[name].Text = default_value
        self._refresh_csv_files()
        print "File paths reset to their original values."

    # --- "Files" tab: list of available CSV files ---

    def _refresh_csv_files(self):
        """
        Does: rebuilds panelCsvFiles (a list, not tiles) from the current contents of CSV_EXPORT_FOLDER.
        Depends on: CSV_EXPORT_FOLDER, os.listdir, self._build_csv_row.
        Returns: nothing (side effect on self.panel_csv_files).
        """
        self.panel_csv_files.Children.Clear()
        try:
            names = sorted(f for f in os.listdir(CSV_EXPORT_FOLDER) if f.lower().endswith(".csv"))
        except Exception:
            names = []

        if not names:
            placeholder = TextBlock()
            placeholder.Text = "(No CSV file yet)"
            placeholder.Foreground = SEARCH_PLACEHOLDER_BRUSH
            placeholder.Margin = Thickness(6)
            self.panel_csv_files.Children.Add(placeholder)
            return

        for name in names:
            self.panel_csv_files.Children.Add(self._build_csv_row(name, os.path.join(CSV_EXPORT_FOLDER, name)))

    def _build_csv_row(self, name, path):
        """
        Does: builds a row of the CSV list (name on the left, Open/Show in folder buttons on the right).
        Depends on: self._make_view_handler/_make_show_in_folder_handler, _shared_resources.
        Returns: Border, the row ready to be added to panelCsvFiles (bottom border = table separator).
        """
        grid = Grid()
        col_name = ColumnDefinition()
        col_name.Width = GridLength(1, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col_name)
        col_actions = ColumnDefinition()
        col_actions.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(col_actions)

        text_block = TextBlock()
        text_block.Text = name
        text_block.VerticalAlignment = VerticalAlignment.Center
        text_block.TextTrimming = TextTrimming.CharacterEllipsis
        text_block.Margin = Thickness(0, 0, 10, 0)
        Grid.SetColumn(text_block, 0)
        grid.Children.Add(text_block)

        actions = StackPanel()
        actions.Orientation = Orientation.Horizontal
        actions.VerticalAlignment = VerticalAlignment.Center

        btn_open = _themed_button()
        btn_open.Content = "Open"
        btn_open.Padding = Thickness(8, 3, 8, 3)
        btn_open.FontSize = 11
        btn_open.Margin = Thickness(0, 0, 6, 0)
        btn_open.Click += self._make_view_handler(path)
        actions.Children.Add(btn_open)

        btn_show = _themed_button()
        btn_show.Content = "Show in folder"
        btn_show.Padding = Thickness(8, 3, 8, 3)
        btn_show.FontSize = 11
        btn_show.Click += self._make_show_in_folder_handler(path)
        actions.Children.Add(btn_show)

        Grid.SetColumn(actions, 1)
        grid.Children.Add(actions)

        row = Border()
        row.BorderBrush = _shared_resources["CardBorderBrush"]
        row.BorderThickness = Thickness(0, 0, 0, 1)
        row.Padding = Thickness(4, 6, 4, 6)
        row.Child = grid
        return row

    def _make_view_handler(self, path):
        """
        Does: closes over path by value to produce the Click handler of the "Open" button of a CSV row.
        Depends on: self._on_view_file.
        Returns: function, the handler(sender, e) to wire on btn_open.Click.
        """
        def handler(sender, e):
            """
            Does: opens the CSV file with its associated application.
            Depends on: self._on_view_file, path captured by the closure.
            Returns: nothing (side effect: launches the associated application).
            """
            self._on_view_file(path)
        return handler

    def _make_show_in_folder_handler(self, path):
        """
        Does: closes over path by value to produce the Click handler of the "Show in folder" button.
        Depends on: self._on_show_in_folder.
        Returns: function, the handler(sender, e) to wire on btn_show.Click.
        """
        def handler(sender, e):
            """
            Does: opens Windows Explorer with the file highlighted.
            Depends on: self._on_show_in_folder, path captured by the closure.
            Returns: nothing (side effect: launches explorer.exe).
            """
            self._on_show_in_folder(path)
        return handler

    def _on_view_file(self, path):
        """
        Does: opens a file with its associated application (equivalent of a double-click in Explorer).
        Depends on: System.Diagnostics.Process.Start.
        Returns: nothing (side effect: launches the associated application, shows a MessageBox on failure).
        """
        try:
            Process.Start(path)
        except Exception as ex:
            MessageBox.Show("Unable to open the file:\n" + str(ex),
                             "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_show_in_folder(self, path):
        """
        Does: opens Windows Explorer with the file already selected/highlighted ("Show in folder").
        Depends on: System.Diagnostics.Process.Start("explorer.exe", "/select,...") (.NET API).
        Returns: nothing (side effect: launches explorer.exe, shows a MessageBox on failure).
        """
        try:
            Process.Start("explorer.exe", '/select,"{}"'.format(path))
        except Exception as ex:
            MessageBox.Show("Unable to show the file in its folder:\n" + str(ex),
                             "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    # --- "Files" tab: cleaning up data folders ---
    # One tile per DATA_ROOT subfolder (excluding legends, see list_data_cleanup_folders in
    # 00_constants.py): size + file count, individual "Clear" button (light red), and a
    # global "Delete all" button (bright red, btnDeleteAllData) that empties all these folders at
    # once. Legends are never affected: they are configuration files reused
    # from one generation to the next, not disposable exports.

    def _refresh_data_cleanup_tiles(self):
        """
        Does: rebuilds panelDataCleanup from the current contents of DATA_ROOT (excluding legends).
        Depends on: list_data_cleanup_folders (00_constants.py), self._build_data_cleanup_tile.
        Returns: nothing (side effect on self.panel_data_cleanup).
        """
        self.panel_data_cleanup.Children.Clear()
        folders = list_data_cleanup_folders()

        if not folders:
            placeholder = TextBlock()
            placeholder.Text = "(No data folder yet)"
            placeholder.Foreground = SEARCH_PLACEHOLDER_BRUSH
            placeholder.Margin = Thickness(6)
            self.panel_data_cleanup.Children.Add(placeholder)
            return

        for name, path in folders:
            self.panel_data_cleanup.Children.Add(self._build_data_cleanup_tile(name, path))

    def _build_data_cleanup_tile(self, name, path):
        """
        Does: builds a cleanup tile (folder name, size, file count, "Clear" button).
        Depends on: get_folder_stats/format_folder_size (00_constants.py), self._make_clear_folder_handler, _shared_resources, CARD_NORMAL_BACKGROUND.
        Returns: Border, the tile ready to be added to panelDataCleanup (stretchable, panelDataCleanup is a 2x2 UniformGrid).
        """
        size_bytes, file_count = get_folder_stats(path)

        content = StackPanel()

        title = TextBlock()
        title.Text = name
        title.FontWeight = FontWeights.SemiBold
        title.TextTrimming = TextTrimming.CharacterEllipsis
        content.Children.Add(title)

        detail = TextBlock()
        detail.Text = "{} - {} file(s)".format(format_folder_size(size_bytes), file_count)
        detail.FontSize = 10
        detail.Foreground = SEARCH_PLACEHOLDER_BRUSH
        detail.Margin = Thickness(0, 2, 0, 6)
        content.Children.Add(detail)

        btn_clear = Button()
        btn_clear.Content = "Clear"
        btn_clear.Style = _shared_resources["DangerButtonLight"]
        btn_clear.Padding = Thickness(8, 3, 8, 3)
        btn_clear.FontSize = 11
        btn_clear.HorizontalAlignment = HorizontalAlignment.Left
        btn_clear.Click += self._make_clear_folder_handler(name, path)
        content.Children.Add(btn_clear)

        # No fixed Width (unlike the app's other tiles): panelDataCleanup is a
        # 2x2 UniformGrid (see XAML), each cell must stretch to fill its allotted space.
        tile = Border()
        tile.Background = CARD_NORMAL_BACKGROUND
        tile.BorderBrush = _shared_resources["CardBorderBrush"]
        tile.BorderThickness = Thickness(1)
        tile.CornerRadius = CornerRadius(0)
        tile.Padding = Thickness(10)
        tile.Margin = Thickness(4)
        tile.Child = content

        return tile

    def _make_clear_folder_handler(self, name, path):
        """
        Does: closes over name/path by value to produce the Click handler of a cleanup tile's "Clear" button.
        Depends on: self._on_clear_folder.
        Returns: function, the handler(sender, e) to wire on btn_clear.Click.
        """
        def handler(sender, e):
            self._on_clear_folder(name, path)
        return handler

    def _on_clear_folder(self, name, path):
        """
        Does: asks for confirmation then clears a data folder (a tile's "Clear" button).
        Depends on: clear_folder_contents (00_constants.py), REPORT_OUTPUT_FOLDER, self._reset_report_status_tile, self._refresh_csv_files/_refresh_data_cleanup_tiles, MessageBox.
        Returns: nothing (side effect on the file system and the UI, if confirmed).
        """
        answer = MessageBox.Show(
            "Delete all contents of \"{}\"? This action cannot be undone.".format(name),
            "Confirm deletion", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        if answer != MessageBoxResult.Yes:
            return

        clear_folder_contents(path)
        print "Folder cleared: " + path

        # The last generated report may no longer exist if this is precisely the folder that was
        # just cleared: the "report result" tile (Open/Show in folder) must revert to the neutral state.
        if os.path.normcase(os.path.normpath(path)) == os.path.normcase(os.path.normpath(REPORT_OUTPUT_FOLDER)):
            self._reset_report_status_tile()

        self._refresh_csv_files()
        self._refresh_data_cleanup_tiles()

    def _reset_report_status_tile(self):
        """
        Does: resets the "report result" tile to the neutral state (e.g. after deleting the reports folder).
        Depends on: _shared_resources, self.border_report_tile/lbl_report_name/btn_report_view/btn_report_show_in_folder.
        Returns: nothing (side effect on self._last_report_path and the report tile's controls).
        """
        self._last_report_path = None
        self.border_report_tile.Background = _shared_resources["SecondaryBackgroundBrush"]
        self.lbl_report_name.Text = "No report generated"
        self.btn_report_view.IsEnabled = False
        self.btn_report_show_in_folder.IsEnabled = False

    def _on_delete_all_data(self, sender, e):
        """
        Does: asks for confirmation then clears ALL data folders (excluding legends) - "Delete all" button.
        Depends on: list_data_cleanup_folders/clear_folder_contents (00_constants.py), self._reset_report_status_tile, self._refresh_csv_files/_refresh_data_cleanup_tiles, MessageBox.
        Returns: nothing (side effect on the file system and the UI, if confirmed).
        """
        folders = list_data_cleanup_folders()
        if not folders:
            return

        answer = MessageBox.Show(
            "Delete all contents of {} data folder(s) (images, CSV, 3D exports, reports...)? "
            "Legends are not affected. This action cannot be undone.".format(len(folders)),
            "Confirm global deletion", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        if answer != MessageBoxResult.Yes:
            return

        for _name, path in folders:
            clear_folder_contents(path)
        print "{} data folder(s) cleared (legends kept).".format(len(folders))

        self._reset_report_status_tile()
        self._refresh_csv_files()
        self._refresh_data_cleanup_tiles()

    # --- "Files" tab: progress + generated report ---

    def _reset_generation_ui(self, total):
        """
        Does: resets the status tile to the neutral state at the start of generation.
        Depends on: _shared_resources, the progress_fill/lbl_progress_status/border_*/btn_report_* controls.
        Returns: nothing (side effect on the status tile's controls).
        """
        self.progress_fill.Width = 0
        self.lbl_progress_status.Text = "Generating... (0/{})".format(total)
        self.border_report_status.Background = _shared_resources["CardBackgroundBrush"]
        self.border_report_tile.Background = _shared_resources["SecondaryBackgroundBrush"]
        self.lbl_report_name.Text = "Generating..."
        self.btn_report_view.IsEnabled = False
        self.btn_report_show_in_folder.IsEnabled = False

    def _update_generation_progress(self, done, total):
        """
        Does: advances the progress bar, refreshes the CSV grid and cleanup tiles, pumps the Win32 message loop.
        Depends on: self.progress_fill/progress_track/lbl_progress_status, self._refresh_csv_files/_refresh_data_cleanup_tiles, SWF.Application.DoEvents().
        Returns: nothing (side effect on the UI; DoEvents() keeps the window responsive during generation).
        """
        # SWF.Application.DoEvents() (same technique as _set_result_display_time in
        # 05_interactive_slides.py): without this call, the window freezes until generation finishes.
        fraction = float(done) / total if total else 1.0
        self.progress_fill.Width = fraction * self.progress_track.ActualWidth
        self.lbl_progress_status.Text = "Generating... ({}/{})".format(done, total)
        self._refresh_csv_files()
        self._refresh_data_cleanup_tiles()
        SWF.Application.DoEvents()

    def _mark_report_ready(self, path):
        """
        Does: activates the PPTX tile (Open/Show in folder) once the report has been generated.
        Depends on: _shared_resources["ReportReadyBackgroundBrush"], self.border_report_tile/lbl_report_name/btn_report_*.
        Returns: nothing (side effect on the report tile's controls).
        """
        # Only the "result" sub-tile (borderReportTile) turns green - the enclosing tile and
        # the progress sub-tile stay neutral, so only the element that actually gives
        # access to the report is highlighted.
        self.lbl_progress_status.Text = "Report complete"
        self.progress_fill.Width = self.progress_track.ActualWidth
        self.border_report_tile.Background = _shared_resources["ReportReadyBackgroundBrush"]
        self.lbl_report_name.Text = os.path.basename(path)
        self.btn_report_view.IsEnabled = True
        self.btn_report_show_in_folder.IsEnabled = True

    def _on_view_report(self, sender, e):
        """
        Does: opens the last generated PPTX report (report tile's "Open" button).
        Depends on: self._last_report_path, self._on_view_file.
        Returns: nothing (side effect: launches PowerPoint on the file if a report exists).
        """
        if self._last_report_path:
            self._on_view_file(self._last_report_path)

    def _on_show_report_in_folder(self, sender, e):
        """
        Does: opens Windows Explorer on the last generated PPTX report (report tile's "Show in folder" button).
        Depends on: self._last_report_path, self._on_show_in_folder.
        Returns: nothing (side effect: launches explorer.exe if a report exists).
        """
        if self._last_report_path:
            self._on_show_in_folder(self._last_report_path)

    # --- Building the 9 selection sections ---

    def _build_sections(self):
        """
        Does: populates the selection StackPanels (x:Name panelXxx from the XAML) with one row per Mechanical object.
        Depends on: self._bodies/_bcs/_bolt_pretensions/etc., self._build_section, self._wire_search_box.
        Returns: nothing (side effect: fills self._sections/_section_order and the WPF panels).
        """
        # The last element of each tuple (`tagged`) indicates whether `objects` is a plain list
        # (row_config_factory(obj)) or a list of (obj, analysis) tuples (row_config_factory(obj,
        # analysis), multi-analysis categories - see ReportGeneratorApp.__init__).
        # panel_kind identifies the state of the global side panel to display for this section's
        # "..." (see ReportGeneratorApp._open_config_panel) - None if the category has nothing to
        # configure (Contacts to display: just a checkbox, no view/section/etc).
        section_defs = [
            ("GeometryParts", "panelGeometryParts", "searchGeometryParts", "general",
             "Parts to isolate (geometry)", self._bodies,
             GeometryPartRowConfig, build_geometry_row_display_name, "geometry_part", False),
            ("MeshParts", "panelMeshParts", "searchMeshParts", "general",
             "Mesh part to isolate", self._bodies,
             MeshPartRowConfig, build_mesh_part_row_display_name, "mesh_part", False),
            ("AnalysisContext", "panelAnalysisContext", "searchAnalysisContext", "general",
             "Analysis context (steps, settings)", self._analyses,
             AnalysisContextRowConfig, build_analysis_context_row_display_name, "mesh_part", False),
            ("BoundaryConditions", "panelBoundaryConditions", "searchBoundaryConditions", "conditions",
             "Boundary Conditions", self._bcs,
             SlideRowConfig, build_row_display_name, "result", True),
            ("BoltPretension", "panelBoltPretension", "searchBoltPretension", "conditions",
             "Bolt Pretension", self._bolt_pretensions,
             SlideRowConfig, build_row_display_name, "result", True),
            ("Contacts", "panelContacts", "searchContacts", "conditions",
             "Contacts to display", self._contact_regions,
             ContactRowConfig, build_contact_row_display_name, None, False),
            ("ContactToolConnections", "panelContactToolConnections", "searchContactToolConnections", "conditions",
             "Connection: Contact Tool", self._contact_tool_connections_results,
             SlideRowConfig, build_row_display_name, "result", False),
            ("SolutionInfo", "panelSolutionInfo", "searchSolutionInfo", "conditions",
             "Solution Information", self._solution_info_trackers,
             SolutionInfoRowConfig, build_solution_info_row_display_name, "solution_info", True),
            ("ContactTool", "panelContactTool", "searchContactTool", "results",
             "Contact Tool Results", self._contact_tool_results,
             SlideRowConfig, build_row_display_name, "result", True),
            ("Results", "panelResults", "searchResults", "results",
             "Results", self._results,
             SlideRowConfig, build_row_display_name, "result", True),
            ("BoltTool", "panelBoltTool", "searchBoltTool", "results",
             "Bolt Tool", self._bolt_tool_results,
             SlideRowConfig, build_row_display_name, "result", True),
        ]

        for (name, panel_name, search_name, group_key, label_text, objects,
             row_config_factory, display_name_func, panel_kind, tagged) in section_defs:
            panel = self.window.FindName(panel_name)
            search_box = self.window.FindName(search_name)
            self._init_search_placeholder(search_box)
            rows = self._build_section(panel, objects, row_config_factory, display_name_func,
                                        panel_kind, tagged)

            self._sections[name] = {
                "rows": rows,
                "group_key": group_key,
                "label": label_text,
                "search_box": search_box,
                "panel": panel,
                "panel_kind": panel_kind,
            }
            self._section_order.append(name)
            self._wire_search_box(search_box, rows, panel)
            self._attach_list_fade(panel.Parent)

    def _attach_list_fade(self, scroll):
        """
        Does: adds a white fade at the bottom of scroll (OpacityMask on the ScrollViewer itself,
        no separate overlay Border - unnecessary here since these lists always sit on a solid
        white CardBorder background), visible only when the content actually overflows the
        visible height. Same intent as _build_preview_list_container, mechanism adapted since
        these ScrollViewer are defined directly in the XAML (variable height/MaxHeight="210" or
        stretched, unlike the fixed height of the preview cards).
        Depends on: ITEM_LIST_FADE_HEIGHT, scroll.ScrollChanged/ActualHeight/ScrollableHeight.
        Returns: nothing (side effect: scroll.OpacityMask recalculated on every ScrollChanged).
        """
        # Empty space added below the last row, the same height as the fade: without it, the
        # last row ends up exactly under the fade area (or even truncated by the bottom of the
        # ScrollViewer) and becomes unreadable once fully scrolled to the bottom of the list.
        content = scroll.Content
        if content is not None:
            content.Margin = Thickness(0, 0, 0, ITEM_LIST_FADE_HEIGHT)

        def update_fade(sender, e):
            """
            Does: recalculates scroll's opacity mask on every scroll/content change.
            Depends on: scroll (captured by the closure), ITEM_LIST_FADE_HEIGHT.
            Returns: nothing (side effect on scroll.OpacityMask).
            """
            height = scroll.ActualHeight
            if height <= 0 or scroll.ScrollableHeight <= 0:
                scroll.OpacityMask = None
                return
            fade_start = 1.0 - min(ITEM_LIST_FADE_HEIGHT / height, 0.5)
            fade_mid = fade_start + 0.6 * (1.0 - fade_start)
            brush = LinearGradientBrush()
            brush.StartPoint = Point(0, 0)
            brush.EndPoint = Point(0, 1)
            brush.GradientStops.Add(GradientStop(WpfColor.FromArgb(255, 255, 255, 255), 0))
            brush.GradientStops.Add(GradientStop(WpfColor.FromArgb(255, 255, 255, 255), fade_start))
            brush.GradientStops.Add(GradientStop(WpfColor.FromArgb(200, 255, 255, 255), fade_mid))
            brush.GradientStops.Add(GradientStop(WpfColor.FromArgb(0, 255, 255, 255), 1))
            scroll.OpacityMask = brush
        scroll.ScrollChanged += update_fade

    def _build_section(self, panel, objects, row_config_factory, display_name_func, panel_kind, tagged=False):
        """
        Does: builds all the WPF rows of a section from the given Mechanical objects.
        Depends on: row_config_factory, self._build_row.
        Returns: list of SectionRow, one per object, in the order of `objects`.
        """
        rows = []
        for entry in objects:
            if tagged:
                obj, analysis = entry
                row_config = row_config_factory(obj, analysis)
            else:
                row_config = row_config_factory(entry)
            row = self._build_row(row_config, display_name_func, panel_kind)
            panel.Children.Add(row.border)
            rows.append(row)
        return rows

    def _build_row(self, row_config, display_name_func, panel_kind):
        """
        Does: builds a WPF row (Border > Grid[CheckBox | TextBlock | Button?]) for a row_config.
        Depends on: SectionRow, _row_status_brush, self._make_toggle_handler/_make_config_click_handler.
        Returns: SectionRow, the built and wired row (Checked/Unchecked/Click events).
        """
        # Name column with "Star" width and TextTrimming.CharacterEllipsis: the text can
        # never overflow onto the "..." button or a neighboring control (unlike
        # fixed-pixel columns).
        grid = Grid()

        col_check = ColumnDefinition()
        col_check.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(col_check)

        col_name = ColumnDefinition()
        col_name.Width = GridLength(1, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col_name)

        checkbox = CheckBox()
        checkbox.Margin = Thickness(0)
        checkbox.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(checkbox, 0)
        grid.Children.Add(checkbox)

        text_block = TextBlock()
        text_block.Text = display_name_func(row_config)
        text_block.VerticalAlignment = VerticalAlignment.Center
        text_block.TextTrimming = TextTrimming.CharacterEllipsis
        text_block.Margin = Thickness(6, 0, 6, 0)
        Grid.SetColumn(text_block, 1)
        grid.Children.Add(text_block)

        config_button = None
        if panel_kind:
            col_button = ColumnDefinition()
            col_button.Width = GridLength.Auto
            grid.ColumnDefinitions.Add(col_button)

            config_button = Button()
            # SecondaryButton is a named resource (not a default TargetType-based style):
            # without this line, this button keeps the default light Windows chrome.
            config_button.Style = self.window.FindResource("SecondaryButton")
            config_button.Content = "..."
            config_button.Padding = Thickness(6, 2, 6, 2)
            config_button.MinWidth = 32
            Grid.SetColumn(config_button, 2)
            grid.Children.Add(config_button)

        border = Border()
        border.Padding = Thickness(4, 2, 4, 2)
        border.Margin = Thickness(0, 1, 0, 1)
        border.CornerRadius = CornerRadius(0)
        border.BorderThickness = Thickness(0)
        border.Child = grid

        row = SectionRow(border, checkbox, text_block, config_button, row_config,
                          display_name_func, panel_kind)
        row.border.Background = _row_status_brush(row)

        toggle_handler = self._make_toggle_handler(row)
        checkbox.Checked += toggle_handler
        checkbox.Unchecked += toggle_handler
        if config_button:
            config_button.Click += self._make_config_click_handler(row)

        return row

    def _make_toggle_handler(self, row):
        """
        Does: closes over row by value to produce the Checked/Unchecked handler of a row's CheckBox.
        Depends on: _row_status_brush, self._update_preview.
        Returns: function, the handler(sender, e) to wire on checkbox.Checked/Unchecked.
        """
        def handler(sender, e):
            """
            Does: recalculates the row's status color (3 states) and refreshes the preview.
            Depends on: _row_status_brush, self._update_preview, row captured by the closure.
            Returns: nothing (side effect on row.border.Background and the preview).
            """
            row.border.Background = _row_status_brush(row)
            self._update_preview()
        return handler

    def _make_config_click_handler(self, row):
        """
        Does: closes over row by value to produce the Click handler of a row's "..." button.
        Depends on: self._on_row_config_click.
        Returns: function, the handler(sender, e) to wire on config_button.Click.
        """
        # Closure necessary: the calling loop (_build_section) reuses its loop
        # variable, a direct reference to row would always capture the last iteration.
        def handler(sender, e):
            """
            Does: opens the row's "..." dialog.
            Depends on: self._on_row_config_click, row captured by the closure.
            Returns: nothing (side effect: may modify row.row_config).
            """
            self._on_row_config_click(row)
        return handler

    def _on_row_config_click(self, row):
        """
        Does: opens the global configuration side panel for the clicked row.
        Depends on: row.panel_kind/row_config/display_name_func, self._open_config_panel.
        Returns: nothing (side effect: shows borderConfigPanel).
        """
        def refresh():
            """
            Does: refreshes the row's text/status color and the preview after "Apply".
            Depends on: row (captured by the closure), _row_status_brush, self._update_preview.
            Returns: nothing (side effect on row.text_block/row.border and the preview).
            """
            row.text_block.Text = row.display_name_func(row.row_config)
            row.border.Background = _row_status_brush(row)
            self._update_preview()
        self._open_config_panel(row.panel_kind, row.row_config, refresh)

    # --- Global configuration side panel ("...") ---
    # Replaces the original 4 modal dialog boxes (RowConfigWindow, GeometryPartConfigWindow,
    # MeshPartConfigWindow, SolutionInfoConfigWindow - see SECTION 4/5/5bis/6): a single panel,
    # hidden by default (borderConfigPanel.Visibility = Collapsed), which displays one of 4 field
    # "kinds" depending on the clicked row ("result"/"geometry_part"/"mesh_part"/"solution_info").
    # "Apply" validates and closes (like the old OK button); "Cancel"/the "x" button close without
    # validating (like the old Cancel button/closing the window).

    def _open_config_panel(self, kind, row_config, refresh_callback, bulk_rows=None):
        """
        Does: opens the global side panel for row_config, in the kind state, and shows it. If
        bulk_rows is provided (see self._on_bulk_config_click), the panel opens in bulk mode:
        row_config then only serves to populate the fields' initial values (those of the
        first checked row), and "Apply" writes to ALL rows in bulk_rows (see
        self._on_config_panel_apply) instead of row_config alone.
        Depends on: self.panel_config_panel/border_config_panel, _build_row_config_fields/_build_steps_section_fields/
            _build_geometry_part_fields/_build_mesh_part_fields/_build_solution_info_fields, _ConfigFieldsHolder.
        Returns: nothing (side effect: populates panelConfigPanel, makes borderConfigPanel visible).
        """
        self._config_panel_kind = kind
        self._config_panel_row_config = row_config
        self._config_panel_refresh = refresh_callback
        self._config_panel_fields = _ConfigFieldsHolder()
        self._config_panel_bulk_rows = bulk_rows

        panel = self.panel_config_panel
        panel.Children.Clear()

        lbl_kicker = TextBlock()
        lbl_kicker.Text = "SETTINGS"
        lbl_kicker.FontSize = 11
        lbl_kicker.FontWeight = FontWeights.Bold
        lbl_kicker.Foreground = _shared_resources["TextMutedBrush"]
        lbl_kicker.Margin = Thickness(0, 0, 0, 2)
        panel.Children.Add(lbl_kicker)

        header = Grid()
        col_title = ColumnDefinition()
        col_title.Width = GridLength(1, GridUnitType.Star)
        header.ColumnDefinitions.Add(col_title)
        col_close = ColumnDefinition()
        col_close.Width = GridLength.Auto
        header.ColumnDefinitions.Add(col_close)

        lbl_title = TextBlock()
        lbl_title.Text = ("Bulk configuration ({} items)".format(len(bulk_rows)) if bulk_rows
                           else row_config.obj.Name)
        lbl_title.Style = _shared_resources["CardTitle"]
        lbl_title.TextWrapping = TextWrapping.Wrap
        Grid.SetColumn(lbl_title, 0)
        header.Children.Add(lbl_title)

        btn_close = Button()
        btn_close.Content = _build_close_icon()
        btn_close.Width = 24
        btn_close.Height = 24
        btn_close.Padding = Thickness(0)
        btn_close.Style = _shared_resources["SecondaryButton"]
        btn_close.VerticalAlignment = VerticalAlignment.Top
        btn_close.Click += self._on_config_panel_close
        Grid.SetColumn(btn_close, 1)
        header.Children.Add(btn_close)

        panel.Children.Add(header)

        if bulk_rows:
            # Warning banner in place of the configured/to configure badge (which wouldn't make
            # sense for a group of rows potentially in different states): explicitly reminds
            # the user of the effect of the "Apply" button before it overwrites anything -
            # the only real safeguard of this mode, along with the "Cancel" button always being available.
            banner = Border()
            banner.Background = ROW_STATUS_SELECTED_BRUSH
            banner.CornerRadius = CornerRadius(0)
            banner.Padding = Thickness(8, 6, 8, 6)
            banner.Margin = Thickness(0, 4, 0, 10)
            banner_text = TextBlock()
            banner_text.Text = ("\"Apply\" below will overwrite the settings of all {} checked "
                                 "line(s) in this section.").format(len(bulk_rows))
            banner_text.FontSize = 11
            banner_text.FontWeight = FontWeights.SemiBold
            banner_text.TextWrapping = TextWrapping.Wrap
            banner.Child = banner_text
            panel.Children.Add(banner)
        else:
            badge = Border()
            badge.Background = ROW_STATUS_CONFIGURED_BRUSH if row_config.configured else ROW_STATUS_SELECTED_BRUSH
            badge.CornerRadius = CornerRadius(0)
            badge.Padding = Thickness(6, 2, 6, 2)
            badge.Margin = Thickness(0, 4, 0, 10)
            badge.HorizontalAlignment = HorizontalAlignment.Left
            badge_text = TextBlock()
            badge_text.Text = "configured" if row_config.configured else "to configure"
            badge_text.FontSize = 10
            badge_text.FontWeight = FontWeights.SemiBold
            badge.Child = badge_text
            panel.Children.Add(badge)

        fields = self._config_panel_fields
        if kind == "result":
            if bulk_rows:
                # The number of proposed loadcases is the MINIMUM across all checked rows
                # (they may come from different analyses - BC/BP/Results/etc are
                # multi-analysis, see section_defs): it's impossible to check a step that doesn't
                # exist for one of the rows, rather than silently filtering it out on apply.
                step_count = min(
                    (get_step_count(row.row_config.analysis) if row.row_config.analysis is not None
                     else self._step_count)
                    for row in bulk_rows
                )
            else:
                step_count = (get_step_count(row_config.analysis) if row_config.analysis is not None
                              else self._step_count)
            _build_row_config_fields(fields, panel, row_config, self._views, self._section_plane_labels,
                                      self._legend_names)
            _build_steps_section_fields(fields, panel, row_config, step_count)
        elif kind == "geometry_part":
            _build_geometry_part_fields(fields, panel, row_config, self._views, self._section_plane_labels)
        elif kind == "mesh_part":
            _build_mesh_part_fields(fields, panel, row_config, self._views)
            # The mesh table choice only makes sense for the Mesh row itself
            # (the other "mesh_part" rows - isolated parts, analysis context - don't have one).
            if row_config is self._mesh_view_config:
                panel.Children.Add(_make_field_label("Mesh table:"))
                fields.cmb_mesh_table = ComboBox()
                fields.cmb_mesh_table.Margin = Thickness(0, 4, 0, 12)
                fields.cmb_mesh_table.Items.Add("Default table (ElementSize, Nodes, Elements)")
                fields.cmb_mesh_table.Items.Add("Full table (all properties)")
                fields.cmb_mesh_table.SelectedIndex = 1 if self._mesh_table_full else 0
                panel.Children.Add(fields.cmb_mesh_table)
        elif kind == "solution_info":
            _build_solution_info_fields(fields, panel, row_config)

        buttons = StackPanel()
        buttons.Orientation = Orientation.Horizontal
        buttons.Margin = Thickness(0, 14, 0, 0)

        btn_apply = _themed_button(primary=True)
        btn_apply.Content = "Apply"
        btn_apply.Width = 110
        btn_apply.Margin = Thickness(0, 0, 10, 0)
        btn_apply.Click += self._on_config_panel_apply
        buttons.Children.Add(btn_apply)

        btn_cancel = _themed_button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 100
        btn_cancel.Click += self._on_config_panel_close
        buttons.Children.Add(btn_cancel)

        panel.Children.Add(buttons)

        self.border_config_panel.Visibility = Visibility.Visible

    def _apply_config_fields_to_row_config(self, kind, fields, row_config):
        """
        Does: reads the panel's fields (fields) and applies them to ONE given row_config, based on kind.
        Never calls the Mechanical API (only Python attribute writes on
        row_config): this is what makes this function safe to call in a loop from
        _on_config_panel_apply in bulk mode, unlike rebuilding a WPF panel per
        row (see the history of the previously reverted "bulk config" feature).
        Depends on: _apply_row_config_fields/_apply_steps_section_fields/_apply_geometry_part_fields/
            _apply_mesh_part_fields/_apply_solution_info_fields.
        Returns: nothing (side effect on row_config only).
        """
        if kind == "result":
            _apply_row_config_fields(fields, row_config)
            _apply_steps_section_fields(fields, row_config)
        elif kind == "geometry_part":
            _apply_geometry_part_fields(fields, row_config)
        elif kind == "mesh_part":
            _apply_mesh_part_fields(fields, row_config)
            if row_config is self._mesh_view_config and hasattr(fields, "cmb_mesh_table"):
                self._mesh_table_full = (fields.cmb_mesh_table.SelectedIndex == 1)
        elif kind == "solution_info":
            _apply_solution_info_fields(fields, row_config)
        row_config.configured = True

    def _on_config_panel_apply(self, sender, e):
        """
        Does: validates the configuration currently in the side panel ("Apply" button) and closes
        it. In bulk mode (self._config_panel_bulk_rows non-empty), applies the same fields to
        EVERY checked row (one attribute write per row, without rebuilding a WPF panel
        or calling the Mechanical API in the loop - see _apply_config_fields_to_row_config) then
        refreshes each row and the preview only once at the end.
        Depends on: self._config_panel_kind/_row_config/_fields/_refresh/_bulk_rows,
            _apply_config_fields_to_row_config, _row_status_brush, self._update_preview.
        Returns: nothing (side effect: updates row_config.configured and refreshes the caller, closes the panel).
        """
        kind = self._config_panel_kind
        fields = self._config_panel_fields
        bulk_rows = self._config_panel_bulk_rows

        if bulk_rows:
            for row in bulk_rows:
                self._apply_config_fields_to_row_config(kind, fields, row.row_config)
                row.text_block.Text = row.display_name_func(row.row_config)
                row.border.Background = _row_status_brush(row)
            self._close_config_panel()
            self._update_preview()
            return

        row_config = self._config_panel_row_config
        refresh_callback = self._config_panel_refresh
        self._apply_config_fields_to_row_config(kind, fields, row_config)
        self._close_config_panel()
        if refresh_callback:
            refresh_callback()

    def _on_config_panel_close(self, sender, e):
        """
        Does: closes the side panel without validating ("Cancel" or "x" button).
        Depends on: self._close_config_panel.
        Returns: nothing (side effect: hides borderConfigPanel).
        """
        self._close_config_panel()

    def _close_config_panel(self):
        """
        Does: clears and hides the global configuration side panel.
        Depends on: self.panel_config_panel/border_config_panel.
        Returns: nothing (side effect: resets the _config_panel_* state).
        """
        self._config_panel_kind = None
        self._config_panel_row_config = None
        self._config_panel_fields = None
        self._config_panel_refresh = None
        self._config_panel_bulk_rows = None
        self.panel_config_panel.Children.Clear()
        self.border_config_panel.Visibility = Visibility.Collapsed

    # --- "Combined slide (different results)" tab ---
    # No more separate window (see SECTION 6bis): the template choice populates
    # panelMultiResultTemplateButtons, the grid lives in gridMultiResultCells (8 cells, only the
    # first N of the chosen template are active), and panelMultiResultSidePanel shows one of 3
    # states for the selected cell - no cell (_show_multi_result_placeholder), result choice
    # (_show_multi_result_picker), or full graphics configuration (_show_multi_result_editor,
    # same fields as a normal slide via _build_row_config_fields, without a step notion).

    def _build_multi_result_tab(self):
        """
        Does: populates the template buttons and initializes the grid/side panel when the window loads.
        Depends on: self.panel_multi_result_template_buttons, MULTI_STEP_SLIDE_TEMPLATES, self._set_multi_result_template.
        Returns: nothing (side effect on the "Combined slide" tab).
        """
        self.panel_multi_result_template_buttons.Children.Clear()
        self._mr_template_buttons = {}
        template_counts = sorted(MULTI_STEP_SLIDE_TEMPLATES.keys())
        for count in template_counts:
            btn = _themed_button()
            btn.Content = "{} results".format(count)
            btn.Padding = Thickness(10, 5, 10, 5)
            btn.FontSize = 11
            btn.Margin = Thickness(0, 0, 6, 4)
            btn.Tag = count
            btn.Click += self._on_pick_multi_result_template
            self.panel_multi_result_template_buttons.Children.Add(btn)
            self._mr_template_buttons[count] = btn

        if template_counts:
            self._set_multi_result_template(template_counts[0])
        else:
            self._show_multi_result_placeholder()
            self._update_multi_result_fill_count()

    def _on_pick_multi_result_template(self, sender, e):
        """
        Does: reacts to clicking a template button (2/3/4/6/8 results).
        Depends on: sender.Tag, self._set_multi_result_template.
        Returns: nothing (side effect: switches template, resets the grid under construction).
        """
        self._set_multi_result_template(sender.Tag)

    def _set_multi_result_template(self, count):
        """
        Does: selects a template (number of active cells) and fully resets the grid under construction.
        Depends on: self._mr_cell_configs/_refresh_multi_result_template_buttons/_rebuild_multi_result_grid/_show_multi_result_placeholder/_update_multi_result_fill_count.
        Returns: nothing (side effect on the "Combined slide" tab's state).
        """
        self._mr_template_count = count
        self._mr_cell_configs = [None] * MULTI_RESULT_CELL_TOTAL
        self._refresh_multi_result_template_buttons()
        self._rebuild_multi_result_grid()
        self._show_multi_result_placeholder()
        self._update_multi_result_fill_count()

    def _refresh_multi_result_template_buttons(self):
        """
        Does: highlights (PrimaryButton) the currently selected template's button, the others remaining SecondaryButton.
        Depends on: self._mr_template_buttons/_mr_template_count, _shared_resources.
        Returns: nothing (side effect on the Style of the template buttons).
        """
        for count, btn in self._mr_template_buttons.items():
            btn.Style = _shared_resources["PrimaryButton" if count == self._mr_template_count else "SecondaryButton"]

    def _rebuild_multi_result_grid(self):
        """
        Does: rebuilds the grid's 8 cells (only the first N of the chosen template are active and clickable).
        Depends on: self.grid_multi_result_cells, self._mr_template_count, self._make_multi_result_cell_click_handler.
        Returns: nothing (side effect: repopulates self._mr_cell_borders/_mr_cell_labels and gridMultiResultCells).
        """
        self.grid_multi_result_cells.Children.Clear()
        self._mr_cell_borders = []
        self._mr_cell_labels = []

        for index in range(MULTI_RESULT_CELL_TOTAL):
            active = index < (self._mr_template_count or 0)

            cell = Border()
            cell.BorderBrush = _shared_resources["CardBorderBrush"]
            cell.BorderThickness = Thickness(1)
            cell.CornerRadius = CornerRadius(0)
            cell.Margin = Thickness(3)
            cell.Background = GRID_CELL_UNCONFIGURED_BRUSH if active else GRID_CELL_DISABLED_BRUSH

            label = TextBlock()
            label.TextWrapping = TextWrapping.Wrap
            label.TextTrimming = TextTrimming.CharacterEllipsis
            label.TextAlignment = TextAlignment.Center
            label.HorizontalAlignment = HorizontalAlignment.Center
            label.VerticalAlignment = VerticalAlignment.Center
            label.Foreground = _shared_resources["TextPrimaryBrush"]
            label.FontSize = 11
            label.Margin = Thickness(4)
            cell.Child = label

            if active:
                cell.Cursor = Cursors.Hand
                cell.MouseLeftButtonUp += self._make_multi_result_cell_click_handler(index)

            self.grid_multi_result_cells.Children.Add(cell)
            self._mr_cell_borders.append(cell)
            self._mr_cell_labels.append(label)
            if active:
                self._update_multi_result_cell_visual(index)

        if self._mr_template_count:
            self.lbl_multi_result_hint.Text = (
                "The grid follows the chosen template: only the first {} cells are "
                "configurable. Click a cell to configure it in the panel on the "
                "right - no more separate windows.".format(self._mr_template_count))
        else:
            self.lbl_multi_result_hint.Text = "Choose a template above to get started."

    def _make_multi_result_cell_click_handler(self, index):
        """
        Does: closes over index by value to produce the click handler of an active grid cell.
        Depends on: self._on_multi_result_cell_click.
        Returns: function, the handler(sender, e) to wire on cell.MouseLeftButtonUp.
        """
        def handler(sender, e):
            self._on_multi_result_cell_click(index)
        return handler

    def _on_multi_result_cell_click(self, index):
        """
        Does: selects the clicked cell and shows the appropriate panel on the right (result choice if empty, direct editing if already configured).
        Depends on: self._mr_cell_configs/_mr_selected_cell_index, self._show_multi_result_picker/_show_multi_result_editor, self._update_multi_result_cell_visual.
        Returns: nothing (side effect on the selection state and the side panel).
        """
        self._mr_selected_cell_index = index
        for i in range(len(self._mr_cell_borders)):
            if i < (self._mr_template_count or 0):
                self._update_multi_result_cell_visual(i)

        cfg = self._mr_cell_configs[index]
        if cfg is not None:
            self._show_multi_result_editor(index, cfg)
        else:
            self._show_multi_result_picker(index)

    def _update_multi_result_cell_visual(self, index):
        """
        Does: refreshes the background/text/border of an active cell according to its state (configured, and/or currently selected).
        Depends on: self._mr_cell_configs/_mr_selected_cell_index/_mr_cell_borders/_mr_cell_labels, GRID_CELL_*_BRUSH.
        Returns: nothing (side effect on the cell's WPF controls).
        """
        cfg = self._mr_cell_configs[index]
        border = self._mr_cell_borders[index]
        label = self._mr_cell_labels[index]

        if cfg is not None:
            border.Background = GRID_CELL_CONFIGURED_BRUSH
            label.Text = "Cell {}\n{}\n(current state)".format(index + 1, cfg.obj.Name)
        else:
            border.Background = GRID_CELL_UNCONFIGURED_BRUSH
            label.Text = "Cell {}\nclick to choose a result\n+".format(index + 1)

        if index == self._mr_selected_cell_index:
            border.BorderThickness = Thickness(2)
            border.BorderBrush = GRID_CELL_SELECTED_BORDER_BRUSH
        else:
            border.BorderThickness = Thickness(1)
            border.BorderBrush = _shared_resources["CardBorderBrush"]

    def _show_multi_result_placeholder(self):
        """
        Does: shows the default side panel (no cell selected) and visually deselects the grid.
        Depends on: self.panel_multi_result_side, self._mr_selected_cell_index, self._update_multi_result_cell_visual.
        Returns: nothing (side effect on panelMultiResultSidePanel and the grid).
        """
        self._mr_editing = None
        previous_index = self._mr_selected_cell_index
        self._mr_selected_cell_index = None
        if previous_index is not None and previous_index < len(self._mr_cell_borders):
            self._update_multi_result_cell_visual(previous_index)

        self.panel_multi_result_side.Children.Clear()
        txt = TextBlock()
        txt.Text = ("Choose a template then click a cell in the grid to "
                     "configure it." if self._mr_template_count else
                     "Choose a combined slide template to get started.")
        txt.TextWrapping = TextWrapping.Wrap
        txt.Foreground = _shared_resources["TextMutedBrush"]
        txt.Margin = Thickness(4)
        self.panel_multi_result_side.Children.Add(txt)

    def _show_multi_result_picker(self, index, current_result=None):
        """
        Does: shows in the side panel the (filterable) list of results available for cell index.
        Depends on: self.panel_multi_result_side, self._results, self._init_search_placeholder, self._make_multi_result_pick_handler.
        Returns: nothing (side effect on panelMultiResultSidePanel; repopulates self._mr_picker_rows).
        """
        self._mr_editing = None
        self.panel_multi_result_side.Children.Clear()

        # Same header (kicker + title + "x" in the top right) as _show_multi_result_editor and
        # _open_config_panel: all 3 states of the side panel must close the same way,
        # no text "Close" button at the bottom just for this state.
        header = Grid()
        col_title = ColumnDefinition()
        col_title.Width = GridLength(1, GridUnitType.Star)
        header.ColumnDefinitions.Add(col_title)
        col_close = ColumnDefinition()
        col_close.Width = GridLength.Auto
        header.ColumnDefinitions.Add(col_close)

        title_panel = StackPanel()
        lbl_case = TextBlock()
        lbl_case.Text = "CELL {}".format(index + 1)
        lbl_case.FontWeight = FontWeights.Bold
        lbl_case.FontSize = 11
        lbl_case.Foreground = _shared_resources["TextMutedBrush"]
        title_panel.Children.Add(lbl_case)
        lbl_title = TextBlock()
        lbl_title.Text = "Choose a result"
        lbl_title.Style = _shared_resources["CardTitle"]
        title_panel.Children.Add(lbl_title)
        Grid.SetColumn(title_panel, 0)
        header.Children.Add(title_panel)

        btn_close = Button()
        btn_close.Content = _build_close_icon()
        btn_close.Width = 24
        btn_close.Height = 24
        btn_close.Padding = Thickness(0)
        btn_close.Style = _shared_resources["SecondaryButton"]
        btn_close.VerticalAlignment = VerticalAlignment.Top
        btn_close.Click += self._on_multi_result_close_editor
        Grid.SetColumn(btn_close, 1)
        header.Children.Add(btn_close)

        self.panel_multi_result_side.Children.Add(header)

        search_box = TextBox()
        search_box.Style = _shared_resources["SearchBox"]
        self._init_search_placeholder(search_box)
        search_box.TextChanged += self._on_multi_result_search_changed
        self.panel_multi_result_side.Children.Add(search_box)

        list_panel = StackPanel()
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Height = 380
        scroll.Content = list_panel
        self.panel_multi_result_side.Children.Add(scroll)

        self._mr_picker_rows = []
        result_objects = [obj for obj, _analysis in self._results]
        for result in result_objects:
            row_border = Border()
            row_border.BorderThickness = Thickness(0, 0, 0, 1)
            row_border.BorderBrush = _shared_resources["CardBorderBrush"]
            row_border.Padding = Thickness(6)
            row_border.Cursor = Cursors.Hand
            row_border.Background = GRID_CELL_CONFIGURED_BRUSH if result == current_result else Brushes.Transparent

            row_text = TextBlock()
            row_text.Text = result.Name
            row_text.FontSize = 11
            row_text.TextWrapping = TextWrapping.Wrap
            row_border.Child = row_text

            row_border.MouseLeftButtonUp += self._make_multi_result_pick_handler(index, result)
            list_panel.Children.Add(row_border)
            self._mr_picker_rows.append((row_border, row_text, result))

    def _on_multi_result_search_changed(self, sender, e):
        """
        Does: live-filters the "Choose a result" list based on the typed text (substring, case-insensitive).
        Depends on: sender (the search TextBox), self._mr_picker_rows, SEARCH_PLACEHOLDER.
        Returns: nothing (side effect: Visibility of the rows in self._mr_picker_rows).
        """
        text = sender.Text
        if text == SEARCH_PLACEHOLDER:
            text = ""
        query = text.strip().lower()
        for row_border, row_text, _result in self._mr_picker_rows:
            visible = (not query) or (query in row_text.Text.lower())
            row_border.Visibility = Visibility.Visible if visible else Visibility.Collapsed

    def _make_multi_result_pick_handler(self, index, result):
        """
        Does: closes over index/result by value to produce the click handler of a "Choose a result" list row.
        Depends on: self._on_multi_result_pick.
        Returns: function, the handler(sender, e) to wire on row_border.MouseLeftButtonUp.
        """
        def handler(sender, e):
            self._on_multi_result_pick(index, result)
        return handler

    def _on_multi_result_pick(self, index, result):
        """
        Does: reacts to choosing a result for cell index and switches the side panel to edit mode.
        Depends on: self._mr_cell_configs, SlideRowConfig, self._show_multi_result_editor.
        Returns: nothing (side effect on panelMultiResultSidePanel).
        """
        existing_cfg = self._mr_cell_configs[index]
        # Reuses the existing config (keeps view/section/legend/etc already chosen) if the same
        # result is reselected; starts from a blank config if the user changes result.
        if existing_cfg is not None and existing_cfg.obj == result:
            cfg = existing_cfg
        else:
            cfg = SlideRowConfig(result)
        self._show_multi_result_editor(index, cfg)

    def _show_multi_result_editor(self, index, cfg):
        """
        Does: shows in the side panel the full graphics configuration (view/section/legend/appearance/scale factor, without steps) of cell index for the result cfg.obj.
        Depends on: self.panel_multi_result_side, _build_row_config_fields, _ConfigFieldsHolder, ROW_STATUS_*_BRUSH.
        Returns: nothing (side effect on panelMultiResultSidePanel; initializes self._mr_editing/_mr_editor_fields).
        """
        self._mr_editing = (index, cfg)
        self.panel_multi_result_side.Children.Clear()

        header = Grid()
        col_title = ColumnDefinition()
        col_title.Width = GridLength(1, GridUnitType.Star)
        header.ColumnDefinitions.Add(col_title)
        col_close = ColumnDefinition()
        col_close.Width = GridLength.Auto
        header.ColumnDefinitions.Add(col_close)

        title_panel = StackPanel()
        lbl_case = TextBlock()
        lbl_case.Text = "CELL {}".format(index + 1)
        lbl_case.FontWeight = FontWeights.Bold
        lbl_case.FontSize = 11
        lbl_case.Foreground = _shared_resources["TextMutedBrush"]
        title_panel.Children.Add(lbl_case)
        lbl_result = TextBlock()
        lbl_result.Text = cfg.obj.Name
        lbl_result.Style = _shared_resources["CardTitle"]
        lbl_result.TextWrapping = TextWrapping.Wrap
        title_panel.Children.Add(lbl_result)
        Grid.SetColumn(title_panel, 0)
        header.Children.Add(title_panel)

        btn_close = Button()
        btn_close.Content = _build_close_icon()
        btn_close.Width = 24
        btn_close.Height = 24
        btn_close.Padding = Thickness(0)
        btn_close.Style = _shared_resources["SecondaryButton"]
        btn_close.VerticalAlignment = VerticalAlignment.Top
        btn_close.Click += self._on_multi_result_close_editor
        Grid.SetColumn(btn_close, 1)
        header.Children.Add(btn_close)

        self.panel_multi_result_side.Children.Add(header)

        status_row = StackPanel()
        status_row.Orientation = Orientation.Horizontal
        status_row.Margin = Thickness(0, 4, 0, 10)

        badge = Border()
        badge.Background = ROW_STATUS_CONFIGURED_BRUSH if cfg.configured else ROW_STATUS_SELECTED_BRUSH
        badge.CornerRadius = CornerRadius(0)
        badge.Padding = Thickness(6, 2, 6, 2)
        badge_text = TextBlock()
        badge_text.Text = "configured" if cfg.configured else "to configure"
        badge_text.FontSize = 10
        badge_text.FontWeight = FontWeights.SemiBold
        badge.Child = badge_text
        status_row.Children.Add(badge)

        btn_change = _themed_button()
        btn_change.Content = "Change result"
        btn_change.FontSize = 11
        btn_change.Padding = Thickness(8, 2, 8, 2)
        btn_change.Margin = Thickness(8, 0, 0, 0)
        btn_change.Click += self._make_multi_result_change_handler(index, cfg)
        status_row.Children.Add(btn_change)

        self.panel_multi_result_side.Children.Add(status_row)

        self._mr_editor_fields = _ConfigFieldsHolder()
        _build_row_config_fields(self._mr_editor_fields, self.panel_multi_result_side, cfg,
                                  self._views, self._section_plane_labels, self._legend_names)

        btn_apply = _themed_button(primary=True)
        btn_apply.Content = "Apply"
        btn_apply.Margin = Thickness(0, 10, 0, 0)
        btn_apply.Click += self._on_multi_result_apply
        self.panel_multi_result_side.Children.Add(btn_apply)

    def _make_multi_result_change_handler(self, index, cfg):
        """
        Does: closes over index/cfg by value to produce the handler of the cell editor's "Change result" button.
        Depends on: self._show_multi_result_picker.
        Returns: function, the handler(sender, e) to wire on btn_change.Click.
        """
        def handler(sender, e):
            self._show_multi_result_picker(index, cfg.obj)
        return handler

    def _on_multi_result_close_editor(self, sender, e):
        """
        Does: closes the cell panel (picker or editor) without validating, back to the "no cell selected" state.
        Depends on: self._show_multi_result_placeholder.
        Returns: nothing (side effect on panelMultiResultSidePanel).
        """
        self._show_multi_result_placeholder()

    def _on_multi_result_apply(self, sender, e):
        """
        Does: validates the graphics configuration of the cell currently being edited ("Apply" button).
        Depends on: self._mr_editing/_mr_editor_fields, _apply_row_config_fields, self._update_multi_result_cell_visual/_update_multi_result_fill_count.
        Returns: nothing (side effect: updates self._mr_cell_configs and refreshes the cell/panel).
        """
        if self._mr_editing is None:
            return
        index, cfg = self._mr_editing
        _apply_row_config_fields(self._mr_editor_fields, cfg)
        cfg.configured = True
        self._mr_cell_configs[index] = cfg
        self._update_multi_result_cell_visual(index)
        self._update_multi_result_fill_count()
        self._show_multi_result_editor(index, cfg)

    def _update_multi_result_fill_count(self):
        """
        Does: refreshes the "X / N cells filled" counter above the grid.
        Depends on: self.lbl_multi_result_fill_count, self._mr_template_count/_mr_cell_configs.
        Returns: nothing (side effect on lblMultiResultFillCount).
        """
        if not self._mr_template_count:
            self.lbl_multi_result_fill_count.Text = ""
            return
        filled = sum(1 for cfg in self._mr_cell_configs[:self._mr_template_count] if cfg is not None)
        self.lbl_multi_result_fill_count.Text = "{} / {} cells filled".format(filled, self._mr_template_count)

    def _on_multi_result_add_to_report(self, sender, e):
        """
        Does: validates that all active cells are configured, adds the combined slide to the report preview (deferred generation) and resets the grid to build another one.
        Depends on: self._mr_template_count/_mr_cell_configs, MultiResultSlideConfig, self._multi_result_slides, self._update_preview, self._set_multi_result_template.
        Returns: nothing (side effect: may add an entry to self._multi_result_slides and refresh the preview).
        """
        if not self._mr_template_count:
            MessageBox.Show("First choose a template (number of results to combine).",
                             "No template", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        active_configs = self._mr_cell_configs[:self._mr_template_count]
        missing = active_configs.count(None)
        if missing:
            MessageBox.Show(
                "All cells must be configured before adding the slide to the report "
                "({} cell(s) out of {} missing).".format(missing, self._mr_template_count),
                "Incomplete configuration", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        self._multi_result_slides.append(MultiResultSlideConfig(self._mr_template_count, active_configs))
        self._update_preview()
        print "Combined slide added to the report preview ({} results).".format(self._mr_template_count)

        # Starts over with a blank grid on the same template, to build another one right after.
        self._set_multi_result_template(self._mr_template_count)

    # --- Search fields: greyed-out placeholder text ---

    def _init_search_placeholder(self, search_box):
        """
        Does: initializes a search field with its placeholder text ("Search...", greyed out).
        Depends on: SEARCH_PLACEHOLDER, SEARCH_PLACEHOLDER_BRUSH, SEARCH_TEXT_BRUSH.
        Returns: nothing (side effect: configures search_box and wires GotFocus/LostFocus).
        """
        search_box.Text = SEARCH_PLACEHOLDER
        search_box.Foreground = SEARCH_PLACEHOLDER_BRUSH

        def on_got_focus(sender, e):
            """
            Does: clears the placeholder text when the user clicks into the field.
            Depends on: SEARCH_PLACEHOLDER, SEARCH_TEXT_BRUSH, search_box captured by the closure.
            Returns: nothing (side effect on search_box).
            """
            if search_box.Text == SEARCH_PLACEHOLDER:
                search_box.Text = ""
                search_box.Foreground = SEARCH_TEXT_BRUSH

        def on_lost_focus(sender, e):
            """
            Does: restores the placeholder text if the field is left empty.
            Depends on: SEARCH_PLACEHOLDER, SEARCH_PLACEHOLDER_BRUSH, search_box captured by the closure.
            Returns: nothing (side effect on search_box).
            """
            if not search_box.Text.strip():
                search_box.Text = SEARCH_PLACEHOLDER
                search_box.Foreground = SEARCH_PLACEHOLDER_BRUSH

        search_box.GotFocus += on_got_focus
        search_box.LostFocus += on_lost_focus

    # --- Search within a section ---

    def _wire_search_box(self, search_box, rows, panel):
        """
        Does: wires the Enter key of a search field to trigger the search.
        Depends on: self._perform_search.
        Returns: nothing (side effect: wires search_box.KeyDown).
        """
        def on_key_down(sender, e):
            """
            Does: triggers the search when the user presses Enter.
            Depends on: self._perform_search, search_box/rows/panel captured by the closure.
            Returns: nothing (side effect: sets e.Handled and launches the search).
            """
            if e.Key == Key.Enter:
                e.Handled = True
                self._perform_search(search_box, rows, panel)
        search_box.KeyDown += on_key_down

    def _perform_search(self, search_box, rows, panel):
        """
        Does: finds and selects the next row in rows whose name contains the typed text.
        Depends on: search_box.Text/Tag, rows, panel, SEARCH_HIGHLIGHT_BRUSH, SEARCH_BOX_*_BACKGROUND.
        Returns: nothing (side effect: checks/highlights the found row or colors the field pink if none found).
        """
        # The found row is moved to the very top of panel (VISUAL order only - "rows"
        # keeps its original order, which remains the report generation order). Re-running a
        # search with the same text resumes after the last found occurrence (search_box.Tag
        # stores (text, index), based on rows' original order, unaffected by the move).
        text = search_box.Text
        if text == SEARCH_PLACEHOLDER:
            text = ""
        query = text.strip().lower()
        row_count = len(rows)
        if not query or row_count == 0:
            search_box.Background = SEARCH_BOX_DEFAULT_BACKGROUND
            return

        last_query, last_index = search_box.Tag if search_box.Tag else (None, -1)
        start_index = (last_index + 1) % row_count if last_query == query else 0

        for offset in range(row_count):
            index = (start_index + offset) % row_count
            row = rows[index]
            if query in row.text_block.Text.lower():
                for r in rows:
                    r.border.BorderThickness = Thickness(0)
                row.border.BorderThickness = Thickness(2)
                row.border.BorderBrush = SEARCH_HIGHLIGHT_BRUSH
                row.checkbox.IsChecked = True
                panel.Children.Remove(row.border)
                panel.Children.Insert(0, row.border)
                row.border.BringIntoView()
                search_box.Tag = (query, index)
                search_box.Background = SEARCH_BOX_DEFAULT_BACKGROUND
                return

        search_box.Tag = (query, -1)
        search_box.Background = SEARCH_BOX_NO_MATCH_BACKGROUND

    # --- Live preview ("Report preview" tab) ---
    # One card per THEME (not per row): each (kind, payload) tuple in self._preview_order
    # is either ("general", label) for Geometry/Mesh, or (section_name, None) for a
    # section as soon as AT LEAST ONE of its rows is checked - this single card then groups
    # all the checked items of the section (see _build_preview_card). The order of this list,
    # editable via drag-and-drop, is the one respected when the report is generated (_on_generate).

    # Sections whose card keeps the full per-item detail (view, section, steps, ...); the
    # other sections only display the raw name of each selected item.
    FULL_DETAIL_SECTIONS = ("ContactTool", "ContactToolConnections", "Results", "BoltTool")

    def _collect_desired_preview_entries(self):
        """
        Does: computes the list of entries (kind, payload) that should appear in the preview.
        Depends on: self.chk_geometry/chk_mesh, self._section_order, self._sections[...]["rows"], self._multi_result_slides.
        Returns: list of (kind, payload) tuples, in a natural order (not yet the preview order).
        """
        entries = []
        if self.chk_geometry.IsChecked:
            entries.append(("general", "Geometry"))
        if self.chk_mesh.IsChecked:
            entries.append(("general", "Mesh"))
        for name in self._section_order:
            if any(row.checkbox.IsChecked for row in self._sections[name]["rows"]):
                entries.append((name, None))
        # No checkbox for these entries (added from the "Combined slide" tab, see
        # _on_multi_result_add_to_report): each one is always "desired" until it has been
        # explicitly removed from its card (see _on_delete_multi_result_slide).
        for cfg in self._multi_result_slides:
            entries.append(("MultiResultSlide", cfg))
        return entries

    def _update_preview(self):
        """
        Does: updates self._preview_order based on the checked boxes/rows, without losing the drag-and-drop order.
        Depends on: self._collect_desired_preview_entries, self._render_preview.
        Returns: nothing (side effect on self._preview_order and the preview display).
        """
        # Unchecked entries removed, new ones added at the end, the rest keeps its current position.
        desired = self._collect_desired_preview_entries()
        desired_set = set(desired)

        self._preview_order = [entry for entry in self._preview_order if entry in desired_set]
        kept_set = set(self._preview_order)
        for entry in desired:
            if entry not in kept_set:
                self._preview_order.append(entry)
                kept_set.add(entry)

        self._render_preview()

    def _build_preview_list_row(self, primary_text, secondary_lines=None):
        """
        Does: builds ONE row of a preview card's vertical list (name + any details).
        Depends on: _shared_resources["CardBorderBrush"], SEARCH_PLACEHOLDER_BRUSH.
        Returns: Border, the row ready to be added to the list container (self._build_preview_list_container).
        """
        # Used identically for ALL categories: one row per parameter/item,
        # instead of a shared text block that becomes unreadable as soon as several items each
        # have several parameters (view, section, steps, ...).
        inner = StackPanel()

        primary_block = TextBlock()
        primary_block.Text = primary_text
        primary_block.FontSize = 11
        primary_block.TextWrapping = TextWrapping.Wrap
        inner.Children.Add(primary_block)

        for line in (secondary_lines or []):
            detail_block = TextBlock()
            detail_block.Text = line
            detail_block.FontSize = 9
            detail_block.Foreground = SEARCH_PLACEHOLDER_BRUSH
            detail_block.TextWrapping = TextWrapping.Wrap
            detail_block.Margin = Thickness(0, 1, 0, 0)
            inner.Children.Add(detail_block)

        row = Border()
        row.BorderBrush = _shared_resources["CardBorderBrush"]
        row.BorderThickness = Thickness(0, 0, 0, 1)
        row.Padding = Thickness(4, 4, 4, 4)
        row.Child = inner
        return row

    def _build_preview_list_container(self, rows):
        """
        Does: builds the vertical list container of a preview card (slightly grey background,
        sitting directly under the title bar): FIXED height (PREVIEW_LIST_DEFAULT_HEIGHT, the same
        for all cards) and scrollable, with a fade at the bottom (PREVIEW_LIST_FADE_HEIGHT) visible
        ONLY if the list actually overflows the visible height (checked after layout, see on_list_loaded).
        Depends on: PREVIEW_LIST_DEFAULT_HEIGHT/FADE_HEIGHT/BACKGROUND(_COLOR), rows already built by self._build_preview_list_row.
        Returns: Grid, the container (scrollable list + overlaid fade) ready to be added to the card.
        """
        list_panel = StackPanel()
        # Same reason as in _attach_list_fade: without this space, the last row ends up
        # exactly under the fade area once fully scrolled to the bottom of the list and becomes unreadable.
        list_panel.Margin = Thickness(0, 0, 0, PREVIEW_LIST_FADE_HEIGHT)
        for row in rows:
            list_panel.Children.Add(row)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Height = PREVIEW_LIST_DEFAULT_HEIGHT
        scroll.Content = list_panel

        list_border = Border()
        list_border.Background = PREVIEW_LIST_BACKGROUND
        list_border.BorderBrush = _shared_resources["CardBorderBrush"]
        list_border.BorderThickness = Thickness(1)
        list_border.Child = scroll

        container = Grid()
        container.Children.Add(list_border)

        fade = Border()
        fade.Height = PREVIEW_LIST_FADE_HEIGHT
        fade.VerticalAlignment = VerticalAlignment.Bottom
        fade.IsHitTestVisible = False
        fade.Visibility = Visibility.Collapsed

        # 3 stops rather than 2 (transparent -> opaque halfway through the fade -> opaque): the
        # gradient "holds" its opacity earlier/stronger instead of fading linearly over
        # the whole height - a more pronounced, less washed-out look than a plain linear gradient.
        fade_brush = LinearGradientBrush()
        fade_brush.StartPoint = Point(0, 0)
        fade_brush.EndPoint = Point(0, 1)
        transparent_color = WpfColor.FromArgb(0, PREVIEW_LIST_BACKGROUND_COLOR.R,
                                               PREVIEW_LIST_BACKGROUND_COLOR.G, PREVIEW_LIST_BACKGROUND_COLOR.B)
        mid_color = WpfColor.FromArgb(200, PREVIEW_LIST_BACKGROUND_COLOR.R,
                                       PREVIEW_LIST_BACKGROUND_COLOR.G, PREVIEW_LIST_BACKGROUND_COLOR.B)
        fade_brush.GradientStops.Add(GradientStop(transparent_color, 0))
        fade_brush.GradientStops.Add(GradientStop(mid_color, 0.45))
        fade_brush.GradientStops.Add(GradientStop(PREVIEW_LIST_BACKGROUND_COLOR, 1))
        fade.Background = fade_brush

        def on_list_loaded(sender, e):
            """
            Does: only shows the fade once the content's actual height is known (after layout).
            Depends on: scroll.ScrollableHeight, fade captured by the closure.
            Returns: nothing (side effect on fade.Visibility).
            """
            fade.Visibility = Visibility.Visible if scroll.ScrollableHeight > 0 else Visibility.Collapsed
        scroll.Loaded += on_list_loaded

        container.Children.Add(fade)
        return container

    def _build_preview_card_content(self, title, chips, order_number):
        """
        Does: builds a preview card's content - title+badge bar (white background, inherited
        from the card), sitting directly above the vertical list container of selected
        items (self._build_preview_list_container).
        Depends on: _shared_resources["AccentBrush"], self._build_preview_list_container (via chips already built by self._build_preview_list_row).
        Returns: tuple (StackPanel content, TextBlock badge) - the badge is returned separately so it can be renumbered without rebuilding the card.
        """
        content = StackPanel()

        title_row = StackPanel()
        title_row.Orientation = Orientation.Horizontal
        title_row.Margin = Thickness(0, 0, 0, 6)

        badge = Border()
        badge.Width = 20
        badge.Height = 20
        badge.CornerRadius = CornerRadius(0)
        badge.Background = _shared_resources["AccentBrush"]
        badge.Margin = Thickness(0, 0, 8, 0)

        badge_text = TextBlock()
        badge_text.Text = str(order_number)
        badge_text.Foreground = Brushes.White
        badge_text.FontSize = 11
        badge_text.FontWeight = FontWeights.Bold
        badge_text.HorizontalAlignment = HorizontalAlignment.Center
        badge_text.VerticalAlignment = VerticalAlignment.Center
        badge.Child = badge_text
        title_row.Children.Add(badge)

        title_block = TextBlock()
        title_block.Text = title
        title_block.FontWeight = FontWeights.Bold
        title_block.TextWrapping = TextWrapping.Wrap
        title_block.VerticalAlignment = VerticalAlignment.Center
        title_row.Children.Add(title_block)

        content.Children.Add(title_row)

        # Always added, even if chips is empty (e.g. Geometry/Mesh without a configured view):
        # keeps a uniform card height in all cases (see PREVIEW_LIST_DEFAULT_HEIGHT),
        # rather than an abnormally short card whenever there is nothing to list.
        content.Children.Add(self._build_preview_list_container(chips))

        return content, badge_text

    def _build_preview_card(self, entry, index):
        """
        Does: builds a complete preview card (background, border, shadow) for an entry from self._preview_order.
        Depends on: self._sections, self.FULL_DETAIL_SECTIONS, self._geometry_view_config/_mesh_view_config, self._build_preview_card_content, self._begin_potential_drag, build_row_display_name (05_interactive_slides.py, for MultiResultSlide).
        Returns: Border, the card ready to be added to panelPreview.
        """
        # Only THE CATEGORY CARD itself is draggable/a drop target (see _begin_potential_drag);
        # the inner chips are not.
        kind, payload = entry

        delete_handler = None

        if kind == "general":
            chips = []
            if payload == "Geometry":
                if self._geometry_view_config.view_name:
                    chips.append(self._build_preview_list_row("view=" + self._geometry_view_config.view_name))
            elif payload == "Mesh":
                table_mode = "Full table" if self._mesh_table_full else "Default table"
                chips.append(self._build_preview_list_row(table_mode))
                if self._mesh_view_config.view_name:
                    chips.append(self._build_preview_list_row("view=" + self._mesh_view_config.view_name))
            content, badge = self._build_preview_card_content(payload, chips, index + 1)
        elif kind == "MultiResultSlide":
            # No checkbox for this entry type (see _on_multi_result_add_to_report): the
            # card itself carries a "Delete" button to remove it from the preview/generation.
            chips = []
            for cell_cfg in payload.cell_configs:
                full_text = build_row_display_name(cell_cfg)
                parts = full_text.split(" | ")
                chips.append(self._build_preview_list_row(parts[0], parts[1:]))
            title = "Combined slide ({} results)".format(payload.template_count)
            content, badge = self._build_preview_card_content(title, chips, index + 1)
            delete_handler = self._make_multi_result_delete_handler(payload)
        else:
            section = self._sections[kind]
            checked_rows = [row for row in section["rows"] if row.checkbox.IsChecked]
            if kind in self.FULL_DETAIL_SECTIONS:
                # "Result" categories: one chip per result, with its full detail (view, section, steps, ...).
                chips = []
                for row in checked_rows:
                    full_text = row.display_name_func(row.row_config)
                    parts = full_text.split(" | ")
                    chips.append(self._build_preview_list_row(parts[0], parts[1:]))
            else:
                # "Context" categories (parts, BC, BP, contacts, solution info, analyses): one
                # chip per name, no detail - via display_name_func()[0] (not the raw obj.Name)
                # so the analysis suffix (see analysis_suffix) also shows up here for
                # Bolt Pretension / Solution Information on a multi-analysis project.
                chips = [self._build_preview_list_row(row.display_name_func(row.row_config).split(" | ")[0])
                         for row in checked_rows]
            content, badge = self._build_preview_card_content(section["label"], chips, index + 1)

        if delete_handler is not None:
            btn_delete = _themed_button()
            btn_delete.Content = "Delete"
            btn_delete.FontSize = 10
            btn_delete.Padding = Thickness(6, 1, 6, 1)
            btn_delete.Margin = Thickness(10, 0, 0, 0)
            btn_delete.VerticalAlignment = VerticalAlignment.Center
            btn_delete.Click += delete_handler
            content.Children[0].Children.Add(btn_delete)  # content.Children[0] = title_row (StackPanel horizontal, see _build_preview_card_content)

        self._entry_to_badge[entry] = badge

        card = Border()
        card.Background = CARD_NORMAL_BACKGROUND
        card.BorderBrush = _shared_resources["CardBorderBrush"]
        card.BorderThickness = Thickness(1)
        card.CornerRadius = CornerRadius(0)
        card.Padding = Thickness(10)
        card.Margin = Thickness(4)
        card.Width = CARD_WIDTH
        card.Cursor = Cursors.SizeAll
        card.Tag = entry
        card.Child = content

        def on_mouse_enter(sender, e):
            """
            Does: turns the card very light blue on hover, except during an active drag.
            Depends on: self._drag_active, CARD_HOVER_BACKGROUND, card captured by the closure.
            Returns: nothing (side effect on card.Background).
            """
            if not self._drag_active:
                card.Background = CARD_HOVER_BACKGROUND

        def on_mouse_leave(sender, e):
            """
            Does: restores the card's normal background when the mouse leaves it.
            Depends on: CARD_NORMAL_BACKGROUND, card captured by the closure.
            Returns: nothing (side effect on card.Background).
            """
            card.Background = CARD_NORMAL_BACKGROUND

        def on_preview_mouse_down(sender, e):
            """
            Does: records the starting point of a potential drag on this card, unless the
            click comes from a nested button (e.g. "Delete" of a combined slide) - otherwise
            CaptureMouse() on panelPreview would prevent the button's Click from firing normally.
            Depends on: self._begin_potential_drag, self._is_button_descendant, card/entry captured by the closure.
            Returns: nothing (side effect: initializes the drag state).
            """
            if self._is_button_descendant(e.OriginalSource):
                return
            self._begin_potential_drag(card, entry, e)

        card.MouseEnter += on_mouse_enter
        card.MouseLeave += on_mouse_leave
        card.PreviewMouseLeftButtonDown += on_preview_mouse_down

        return card

    def _render_preview(self):
        """
        Does: fully rebuilds the panelPreview WrapPanel's cards from self._preview_order.
        Depends on: self._build_preview_card, self._preview_order.
        Returns: nothing (side effect on self.panel_preview and self._entry_to_card/_entry_to_badge).
        """
        self.panel_preview.Children.Clear()
        self._entry_to_card = {}
        self._entry_to_badge = {}

        for index, entry in enumerate(self._preview_order):
            card = self._build_preview_card(entry, index)
            self._entry_to_card[entry] = card
            self.panel_preview.Children.Add(card)

        if not self._preview_order:
            placeholder = TextBlock()
            placeholder.Text = "(No slide selected)"
            placeholder.Foreground = SEARCH_PLACEHOLDER_BRUSH
            placeholder.Margin = Thickness(6)
            self.panel_preview.Children.Add(placeholder)

    # --- Preview card drag-and-drop ---
    # The mouse is captured by panelPreview (not by the card itself): that way, reordering
    # the WrapPanel's children during the drag never loses the capture, even if the
    # source card is briefly removed/reinserted. An undecorated Popup ("ghost", a visual
    # copy via VisualBrush) follows the mouse; the other cards shift live as soon as
    # the cursor hovers over another card.

    def _begin_potential_drag(self, card, entry, e):
        """
        Does: records the starting point of a potential drag and captures the mouse on panelPreview.
        Depends on: self.panel_preview.CaptureMouse().
        Returns: nothing (side effect: initializes self._drag_pending_card/_drag_pending_entry/_drag_start_point).
        """
        self._drag_pending_card = card
        self._drag_pending_entry = entry
        self._drag_start_point = e.GetPosition(self.panel_preview)
        self.panel_preview.CaptureMouse()

    def _on_preview_panel_mouse_move(self, sender, e):
        """
        Does: starts the drag once past a movement threshold, then makes the ghost follow and reorders on hover.
        Depends on: self._drag_pending_card/_drag_pending_entry/_drag_start_point, self._start_drag/_update_drag_ghost_position/_update_drag_hover.
        Returns: nothing (side effect: triggers the drag or updates its position).
        """
        if e.LeftButton != MouseButtonState.Pressed or self._drag_pending_entry is None:
            return

        current_point = e.GetPosition(self.panel_preview)

        if not self._drag_active:
            delta_x = abs(current_point.X - self._drag_start_point.X)
            delta_y = abs(current_point.Y - self._drag_start_point.Y)
            if delta_x < 6 and delta_y < 6:
                return
            self._start_drag(self._drag_pending_card, self._drag_pending_entry)

        self._update_drag_ghost_position(e)
        self._update_drag_hover(current_point)

    def _on_preview_panel_mouse_up(self, sender, e):
        """
        Does: releases the mouse capture and cleanly ends the current drag (if any).
        Depends on: self.panel_preview.ReleaseMouseCapture(), self._drag_active, self._end_drag.
        Returns: nothing (side effect: resets the pending drag state).
        """
        self.panel_preview.ReleaseMouseCapture()
        if self._drag_active:
            self._end_drag()
        self._drag_pending_card = None
        self._drag_pending_entry = None
        self._drag_start_point = None

    def _start_drag(self, card, entry):
        """
        Does: actually starts the drag (fades the source card, opens the ghost Popup).
        Depends on: Border/VisualBrush/Popup (WPF), card.ActualWidth/ActualHeight.
        Returns: nothing (side effect: initializes self._drag_active/_drag_entry/_drag_source_card/_drag_popup).
        """
        self._drag_active = True
        self._drag_entry = entry
        self._drag_source_card = card
        card.Opacity = 0.25

        ghost = Border()
        ghost.Width = card.ActualWidth
        ghost.Height = card.ActualHeight
        ghost.Background = VisualBrush(card)
        ghost.Opacity = 0.85

        popup = Popup()
        popup.AllowsTransparency = True
        popup.Placement = PlacementMode.Absolute
        popup.IsHitTestVisible = False
        popup.Focusable = False
        popup.Child = ghost
        popup.IsOpen = True
        self._drag_popup = popup

    def _update_drag_ghost_position(self, e):
        """
        Does: moves the ghost Popup so it stays centered on the cursor.
        Depends on: self._drag_popup, self._drag_source_card, self.panel_preview.PointToScreen.
        Returns: nothing (side effect on self._drag_popup.HorizontalOffset/VerticalOffset).
        """
        screen_point = self.panel_preview.PointToScreen(e.GetPosition(self.panel_preview))
        card = self._drag_source_card
        self._drag_popup.HorizontalOffset = screen_point.X - card.ActualWidth / 2.0
        self._drag_popup.VerticalOffset = screen_point.Y - card.ActualHeight / 2.0

    def _is_button_descendant(self, element):
        """
        Does: determines whether element is a Button or is contained in a Button (e.g. the
        auto-generated TextBlock for Content="Delete"), by walking up the visual tree.
        Depends on: VisualTreeHelper.GetParent (WPF).
        Returns: bool, True as soon as a Button is found on the path.
        """
        node = element
        while node is not None:
            if isinstance(node, Button):
                return True
            node = VisualTreeHelper.GetParent(node)
        return False

    def _find_ancestor_card(self, element):
        """
        Does: walks up the visual tree from a hit-test result to a card's Border.
        Depends on: VisualTreeHelper.GetParent (WPF), the Tag = entry convention on cards (see _build_preview_card).
        Returns: Border or None, the found card (non-None Tag) or None if there isn't one on the path.
        """
        node = element
        while node is not None:
            if isinstance(node, Border) and node.Tag is not None:
                return node
            node = VisualTreeHelper.GetParent(node)
        return None

    def _update_drag_hover(self, position):
        """
        Does: moves the dragged entry within self._preview_order if the cursor hovers over ANOTHER card.
        Depends on: self._find_ancestor_card, self._preview_order, self._reorder_children_to_match_preview_order.
        Returns: nothing (side effect on self._preview_order and the display if a move happens).
        """
        hit = self.panel_preview.InputHitTest(position)
        target_card = self._find_ancestor_card(hit) if hit else None
        if target_card is None or target_card is self._drag_source_card:
            return

        target_entry = target_card.Tag
        if self._drag_entry not in self._preview_order or target_entry not in self._preview_order:
            return

        source_index = self._preview_order.index(self._drag_entry)
        target_index = self._preview_order.index(target_entry)
        if source_index == target_index:
            return

        moved = self._preview_order.pop(source_index)
        self._preview_order.insert(target_index, moved)
        self._reorder_children_to_match_preview_order()

    def _reorder_children_to_match_preview_order(self):
        """
        Does: reorders panelPreview.Children to reflect self._preview_order without recreating the cards.
        Depends on: self._entry_to_card/_entry_to_badge, self._preview_order.
        Returns: nothing (side effect on panelPreview.Children and the numbered badges).
        """
        # Unlike _render_preview, does not recreate the cards: essential during an active
        # drag, to avoid losing the event handlers or the mouse capture (captured
        # on the panel, not on the card). Also renumbers the badges so they stay correct
        # during the drag, not just once released.
        children = self.panel_preview.Children
        for target_index, entry in enumerate(self._preview_order):
            card = self._entry_to_card.get(entry)
            if card is None:
                continue
            current_index = children.IndexOf(card)
            if current_index != target_index:
                children.RemoveAt(current_index)
                children.Insert(target_index, card)

            badge = self._entry_to_badge.get(entry)
            if badge is not None:
                badge.Text = str(target_index + 1)

    def _end_drag(self):
        """
        Does: ends the current drag (closes the ghost, restores the source card's opacity).
        Depends on: self._drag_popup, self._drag_source_card.
        Returns: nothing (side effect: resets the drag state).
        """
        self._drag_active = False
        if self._drag_popup is not None:
            self._drag_popup.IsOpen = False
            self._drag_popup = None
        if self._drag_source_card is not None:
            self._drag_source_card.Opacity = 1.0
        self._drag_source_card = None
        self._drag_entry = None

    def _get_checked_row_configs(self, name):
        """
        Does: retrieves the row_config of checked rows in section 'name'.
        Depends on: self._sections[name]["rows"].
        Returns: list of row_config, those whose CheckBox is checked.
        """
        return [row.row_config for row in self._sections[name]["rows"] if row.checkbox.IsChecked]

    # --- Simple checkbox handlers ---

    def _on_simple_toggle(self, sender, e):
        """
        Does: refreshes the preview when a simple checkbox (Geometry/Mesh) changes state.
        Depends on: self._update_preview.
        Returns: nothing (side effect on the preview).
        """
        self._update_preview()

    def _on_geometry_view_click(self, sender, e):
        """
        Does: opens the global side panel for view selection for the Geometry slide (card's "Settings" button).
        Depends on: self._open_config_panel, self._geometry_view_config, self._refresh_general_slide_status, self._update_preview.
        Returns: nothing (side effect: shows borderConfigPanel; updates self._geometry_view_config.view_name, the card status and the preview on "Apply").
        """
        def refresh():
            self._refresh_general_slide_status()
            self._update_preview()
        self._open_config_panel("mesh_part", self._geometry_view_config, refresh)

    def _on_mesh_view_click(self, sender, e):
        """
        Does: opens the global side panel for view selection for the Mesh slide (card's "Settings" button).
        Depends on: self._open_config_panel, self._mesh_view_config, self._refresh_general_slide_status, self._update_preview.
        Returns: nothing (side effect: shows borderConfigPanel; updates self._mesh_view_config.view_name, the card status and the preview on "Apply").
        """
        def refresh():
            self._refresh_general_slide_status()
            self._update_preview()
        self._open_config_panel("mesh_part", self._mesh_view_config, refresh)

    # --- Select/deselect all, per tab ---

    def _set_group_checked(self, group_key, checked):
        """
        Does: (un)checks every row of every section in a tab (group_key).
        Depends on: self._section_order, self._sections, self._update_preview.
        Returns: nothing (side effect on the CheckBoxes of the affected sections and on the preview).
        """
        for name in self._section_order:
            section = self._sections[name]
            if section["group_key"] != group_key:
                continue
            for row in section["rows"]:
                row.checkbox.IsChecked = checked
        self._update_preview()

    def _set_section_checked(self, name, checked):
        """
        Does: (un)checks every row of a SINGLE section (one specific selection zone).
        Depends on: self._sections[name]["rows"], self._update_preview.
        Returns: nothing (side effect on the section's CheckBoxes and on the preview).
        """
        for row in self._sections[name]["rows"]:
            row.checkbox.IsChecked = checked
        self._update_preview()

    def _make_zone_toggle_handler(self, name, checked):
        """
        Does: closes over name/checked to produce the Click handler for a zone's "Select all"/"Deselect all" button.
        Depends on: self._set_section_checked.
        Returns: function, the handler(sender, e) to wire to the zone's button.
        """
        def handler(sender, e):
            """
            Does: (un)checks every row of the zone associated with the button.
            Depends on: self._set_section_checked, name/checked captured by the closure.
            Returns: nothing (side effect on the zone's CheckBoxes).
            """
            self._set_section_checked(name, checked)
        return handler

    def _make_bulk_config_handler(self, name):
        """
        Does: closes over name to produce the Click handler for a zone's "Configure
        selection..." button.
        Depends on: self._on_bulk_config_click.
        Returns: function, the handler(sender, e) to wire to the zone's button.
        """
        def handler(sender, e):
            """
            Does: opens the bulk configuration panel for the checked rows of the associated zone.
            Depends on: self._on_bulk_config_click, name captured by the closure.
            Returns: nothing (side effect: may open borderConfigPanel).
            """
            self._on_bulk_config_click(name)
        return handler

    def _on_bulk_config_click(self, name):
        """
        Does: opens the global side panel in bulk (group) mode for every checked row of
        zone `name`. Does nothing if the zone has no configuration panel (Contacts) or
        if no row is checked - warns the user in that last case instead of opening
        an empty panel.
        Depends on: self._sections, self._open_config_panel.
        Returns: nothing (side effect: may show a MessageBox or open borderConfigPanel).
        """
        section = self._sections[name]
        panel_kind = section["panel_kind"]
        if not panel_kind:
            return

        checked_rows = [row for row in section["rows"] if row.checkbox.IsChecked]
        if not checked_rows:
            MessageBox.Show(
                "Check at least one line in \"{}\" before configuring the selection.".format(section["label"]),
                "No line selected", MessageBoxButton.OK, MessageBoxImage.Information)
            return

        # A non-None bulk_rows is what switches _open_config_panel/_on_config_panel_apply into
        # group mode (see those methods): the template row_config (first checked row) only
        # serves to populate the initially displayed values, and is never itself written back on
        # "Apply" (each checked row receives its own copy of the chosen values).
        self._open_config_panel(panel_kind, checked_rows[0].row_config, None, bulk_rows=checked_rows)

    def _wire_zone_select_buttons(self):
        """
        Does: wires the "Select all"/"Deselect all"/"Configure selection" buttons of each
        selection zone (XAML card header).
        Depends on: self._section_order, self.window.FindName, self._make_zone_toggle_handler,
            self._make_bulk_config_handler.
        Returns: nothing (side effect: wires the Click of the btnZoneCheck{name}/
            btnZoneUncheck{name}/btnZoneConfig{name} buttons).
        """
        # Per-zone shortcut, in addition to the existing "Select all"/"Deselect all" buttons that
        # act on a whole tab at once (see _set_group_checked).
        for name in self._section_order:
            check_btn = self.window.FindName("btnZoneCheck" + name)
            uncheck_btn = self.window.FindName("btnZoneUncheck" + name)
            config_btn = self.window.FindName("btnZoneConfig" + name)
            if check_btn is not None:
                check_btn.Click += self._make_zone_toggle_handler(name, True)
            if uncheck_btn is not None:
                uncheck_btn.Click += self._make_zone_toggle_handler(name, False)
            # Absent for zones without a configuration panel (Contacts: simple checkbox,
            # nothing to configure - see section_defs in _build_sections).
            if config_btn is not None:
                config_btn.Click += self._make_bulk_config_handler(name)

    def _on_check_all_general(self, sender, e):
        """
        Does: checks every row of the "General slides" tab (including Geometry/Mesh).
        Depends on: self.chk_geometry/chk_mesh, self._set_group_checked.
        Returns: nothing (side effect on the tab's CheckBoxes).
        """
        self.chk_geometry.IsChecked = True
        self.chk_mesh.IsChecked = True
        self._set_group_checked("general", True)

    def _on_uncheck_all_general(self, sender, e):
        """
        Does: unchecks every row of the "General slides" tab (including Geometry/Mesh).
        Depends on: self.chk_geometry/chk_mesh, self._set_group_checked.
        Returns: nothing (side effect on the tab's CheckBoxes).
        """
        self.chk_geometry.IsChecked = False
        self.chk_mesh.IsChecked = False
        self._set_group_checked("general", False)

    def _on_check_all_conditions(self, sender, e):
        """
        Does: checks every row of the "Conditions and contacts" tab.
        Depends on: self._set_group_checked.
        Returns: nothing (side effect on the tab's CheckBoxes).
        """
        self._set_group_checked("conditions", True)

    def _on_uncheck_all_conditions(self, sender, e):
        """
        Does: unchecks every row of the "Conditions and contacts" tab.
        Depends on: self._set_group_checked.
        Returns: nothing (side effect on the tab's CheckBoxes).
        """
        self._set_group_checked("conditions", False)

    def _on_check_all_results(self, sender, e):
        """
        Does: checks every row of the "Result categories" tab.
        Depends on: self._set_group_checked.
        Returns: nothing (side effect on the tab's CheckBoxes).
        """
        self._set_group_checked("results", True)

    def _on_uncheck_all_results(self, sender, e):
        """
        Does: unchecks every row of the "Result categories" tab.
        Depends on: self._set_group_checked.
        Returns: nothing (side effect on the tab's CheckBoxes).
        """
        self._set_group_checked("results", False)

    # --- Utilities: figure deletion / basic view creation ---

    def _on_delete_figures(self, sender, e):
        """
        Does: deletes stale figures generated by previous exports (dedicated button).
        Depends on: remove_stale_figures (05_interactive_slides.py).
        Returns: nothing (side effect: deletes files, shows a MessageBox on failure).
        """
        try:
            remove_stale_figures()
            print "Figures deleted."
        except Exception as ex:
            print "ERROR while deleting figures: " + str(ex)
            MessageBox.Show("Error while deleting figures:\n" + str(ex),
                             "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_reset_legends(self, sender, e):
        """
        Does: resets the result legends applied in Mechanical (dedicated button).
        Depends on: reset_legend (05_interactive_slides.py).
        Returns: nothing (side effect: changes the legends' state, shows a MessageBox on failure).
        """
        try:
            reset_legend()
            print "Legends reset."
        except Exception as ex:
            print "ERROR while resetting legends: " + str(ex)
            MessageBox.Show("Error while resetting legends:\n" + str(ex),
                             "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_create_basic_views(self, sender, e):
        """
        Does: creates the basic views in the View Manager and refreshes the view/section lists.
        Depends on: create_basic_views/collect_views/collect_section_planes/section_plane_label (05_interactive_slides.py).
        Returns: nothing (side effect: creates Mechanical views, updates self._views/_section_planes/_section_plane_labels).
        """
        try:
            created = create_basic_views()
            self._views = collect_views()
            self._section_planes = collect_section_planes()
            self._section_plane_labels = [
                section_plane_label(sp, i) for i, sp in enumerate(self._section_planes)
            ]
            if created:
                print "{} basic view(s) created: {}.".format(len(created), ", ".join(created))
            else:
                print "No basic view could be created (see the Mechanical console)."
        except Exception as ex:
            print "ERROR while creating basic views: " + str(ex)
            MessageBox.Show("Error while creating basic views:\n" + str(ex),
                             "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_export_3d(self, sender, e):
        """
        Does: exports the 3D view of every result and Contact/Bolt Tool (Solution branch) of the project to .avz.
        Depends on: export_all_3d_views, EXPORT_3D_FOLDER (00_constants.py), _print_console_banner.
        Returns: nothing (side effect: creates .avz files in EXPORT_3D_FOLDER, shows a MessageBox on failure).
        """
        try:
            _print_console_banner("3D EXPORT (.avz) IN PROGRESS...")
            exported_count = export_all_3d_views(EXPORT_3D_FOLDER)
            if exported_count:
                _print_console_banner("{} 3D VIEW(S) EXPORTED".format(exported_count))
                print "Available .avz files in: " + EXPORT_3D_FOLDER
            else:
                _print_console_banner("NO 3D VIEW EXPORTED")
                print "No result / Contact Tool / Bolt Tool (Solution branch) found to export."
        except Exception as ex:
            print "ERROR during 3D export: " + str(ex)
            MessageBox.Show("Error during 3D export:\n" + str(ex),
                             "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _make_multi_result_delete_handler(self, cfg):
        """
        Does: closes over cfg to produce the handler for a MultiResultSlide card's "Delete" button.
        Depends on: self._on_delete_multi_result_slide.
        Returns: function, the handler(sender, e) to wire to btn_delete.Click.
        """
        def handler(sender, e):
            self._on_delete_multi_result_slide(cfg)
        return handler

    def _on_delete_multi_result_slide(self, cfg):
        """
        Does: removes a "different results" combined slide from the preview and the generation.
        Depends on: self._multi_result_slides, self._update_preview.
        Returns: nothing (side effect on self._multi_result_slides and the preview).
        """
        if cfg in self._multi_result_slides:
            self._multi_result_slides.remove(cfg)
        self._update_preview()

    def _on_close(self, sender, e):
        """
        Does: closes the application's main window (Close button).
        Depends on: self.window.
        Returns: nothing (side effect: closes self.window).
        """
        self.window.Close()

    # --- Report generation ---

    def _on_generate(self, sender, e):
        """
        Does: generates the PowerPoint report following the order of self._preview_order.
        Depends on: PPTReportBuilder, build_*_slides/create_*_slide (04_slides.py/05_interactive_slides.py), apply_view_if_exists/self._geometry_view_config/_mesh_view_config, self._update_generation_progress.
        Returns: nothing (side effect: creates the PPTX file, updates the status UI, shows a MessageBox on failure).
        """
        # Each card is processed one at a time, in order: a "general" card adds its single
        # slide, a section card adds ALL the slides of its category at once (a batch
        # function from 05_interactive_slides.py). So the reordering granularity (and the
        # progress bar) is the card/category, not the individual slide.
        # PowerPoint stays visible throughout this method (see PPTReportBuilder.__init__ - keeping
        # it invisible turned out to be unstable on a report with many slides); the
        # WPF window stays responsive thanks to SWF.Application.DoEvents() (_update_generation_progress).
        if not self._preview_order:
            MessageBox.Show("No slide selected: check at least one option before generating the report.",
                             "Nothing to generate", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        total = len(self._preview_order)
        report = None
        self._reset_generation_ui(total)

        try:
            remove_stale_figures()

            _print_console_banner("REPORT GENERATION IN PROGRESS...")
            print "Opening PowerPoint template..."
            report = PPTReportBuilder(TEMPLATE_PATH)

            for index, (kind, payload) in enumerate(self._preview_order):
                if kind == "general":
                    if payload == "Geometry":
                        apply_view_if_exists(self._geometry_view_config.view_name, self._views)
                        create_geometry_slide(report)
                    elif payload == "Mesh":
                        apply_view_if_exists(self._mesh_view_config.view_name, self._views)
                        build_mesh_slide(report, self._mesh_table_full)
                    print "Slide added: " + payload
                    self._update_generation_progress(index + 1, total)
                    continue

                if kind == "MultiResultSlide":
                    template = get_multi_step_template(payload.template_count)
                    build_multi_result_slide(report, template, payload.cell_configs, self._views,
                                              self._section_planes, self._section_plane_labels)
                    print "Combined multi-result slide added ({} results).".format(len(payload.cell_configs))
                    self._update_generation_progress(index + 1, total)
                    continue

                selected = self._get_checked_row_configs(kind)
                if not selected:
                    self._update_generation_progress(index + 1, total)
                    continue

                if kind == "AnalysisContext":
                    build_analysis_context_slides(report, selected, self._views)
                elif kind == "GeometryParts":
                    build_geometry_part_slides(report, selected, self._bodies,
                                                self._views, self._section_planes, self._section_plane_labels)
                elif kind == "MeshParts":
                    build_mesh_part_slides(report, selected, self._bodies, self._views)
                elif kind == "BoundaryConditions":
                    build_bc_slides(report, selected, self._views, self._section_planes, self._section_plane_labels)
                elif kind == "BoltPretension":
                    build_bp_slides(report, selected, self._views, self._section_planes, self._section_plane_labels)
                elif kind == "Contacts":
                    build_contact_summary_slide(report, selected)
                elif kind == "SolutionInfo":
                    build_solution_info_slides(report, selected)
                elif kind == "ContactTool":
                    build_result_slides(report, selected, "-- Contact Tool Results --",
                                         self._views, self._section_planes, self._section_plane_labels,
                                         self._analysis)
                elif kind == "ContactToolConnections":
                    build_result_slides(report, selected, "-- Connection: Contact Tool --",
                                         self._views, self._section_planes, self._section_plane_labels,
                                         self._analysis)
                elif kind == "BoltTool":
                    build_result_slides(report, selected, "-- Bolt Tool --",
                                         self._views, self._section_planes, self._section_plane_labels,
                                         self._analysis)
                elif kind == "Results":
                    # Generic subtitle in multi-analysis projects: self._analysis.Name (Analyses[0])
                    # would be misleading for results coming from another analysis - that
                    # information is already shown in each slide's TITLE (see analysis_suffix).
                    results_subtitle = "-- Results --" if self._multi_analysis else self._analysis.Name
                    build_result_slides(report, selected, results_subtitle,
                                         self._views, self._section_planes, self._section_plane_labels,
                                         self._analysis)

                print "{} slide(s) {} added.".format(len(selected), self._sections[kind]["label"])
                self._update_generation_progress(index + 1, total)

            self.btn_generate.IsEnabled = False
            report.keep_open()  # neither Save() nor Close()/Quit() - the report stays open and unsaved in PowerPoint
            self._last_report_path = report.working_copy_path
            self._mark_report_ready(report.working_copy_path)
            _print_console_banner("REPORT GENERATED SUCCESSFULLY")
            print "Report available in the Files tab: " + report.working_copy_path
        except Exception as ex:
            _print_console_banner("ERROR DURING REPORT GENERATION")
            print str(ex)
            if report is not None:
                try:
                    report.close()
                except Exception as close_ex:
                    print "Unable to close PowerPoint: " + str(close_ex)
            MessageBox.Show("Error during report generation:\n" + str(ex),
                             "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    # --- Event wiring ---

    def _wire_events(self):
        """
        Does: wires every event of the main window (buttons, checkboxes, drag-and-drop).
        Depends on: every control found by self._find_controls, the self._on_* handlers.
        Returns: nothing (side effect: subscribes the handlers to the WPF events).
        """
        self.btn_delete_figures.Click += self._on_delete_figures
        self.btn_reset_legends.Click += self._on_reset_legends
        self.btn_create_views.Click += self._on_create_basic_views
        self.btn_export_3d.Click += self._on_export_3d
        self.btn_multi_result_add_to_report.Click += self._on_multi_result_add_to_report

        # The mouse is captured on panelPreview (not on each card) during a drag-and-drop:
        # these two handlers must therefore live here, only once (see _begin_potential_drag).
        self.panel_preview.MouseMove += self._on_preview_panel_mouse_move
        self.panel_preview.PreviewMouseLeftButtonUp += self._on_preview_panel_mouse_up

        self.chk_geometry.Checked += self._on_simple_toggle
        self.chk_geometry.Unchecked += self._on_simple_toggle
        self.btn_geometry_view.Click += self._on_geometry_view_click
        self.chk_mesh.Checked += self._on_simple_toggle
        self.chk_mesh.Unchecked += self._on_simple_toggle
        self.btn_mesh_view.Click += self._on_mesh_view_click

        self.btn_check_all_general.Click += self._on_check_all_general
        self.btn_uncheck_all_general.Click += self._on_uncheck_all_general
        self.btn_check_all_conditions.Click += self._on_check_all_conditions
        self.btn_uncheck_all_conditions.Click += self._on_uncheck_all_conditions
        self.btn_check_all_results.Click += self._on_check_all_results
        self.btn_uncheck_all_results.Click += self._on_uncheck_all_results

        self.btn_generate.Click += self._on_generate
        self.btn_close.Click += self._on_close

        self.btn_report_view.Click += self._on_view_report
        self.btn_report_show_in_folder.Click += self._on_show_report_in_folder


# --- SECTION 8 - Entry point ---
# init/HighFiveOut are the two callbacks declared in "Liebherr Report Generator.xml"
# (<oninit>init</oninit> and <onclick>HighFiveOut</onclick> on the toolbar button). These are the
# ONLY entry points: nothing else in this file runs when the extension loads.


def init():
    """
    Does: oninit callback of the Mechanical interface (Liebherr Report Generator.xml).
    Depends on: nothing - must NOT touch ExtAPI.DataModel.Project.Model here (see _prepare_environment, called later from HighFiveOut).
    Returns: nothing.
    """
    pass


def HighFiveOut(index):
    """
    Does: onclick callback of the toolbar button (Liebherr Report Generator.xml) - the real entry point of the application.
    Depends on: _harden_console_output, _prepare_environment, ReportGeneratorApp.
    Returns: nothing (side effect: displays the WPF report generation window).

    ACT invokes toolbar onclick callbacks with one argument (index of the clicked entry, ACT SDK
    convention) - the signature must accept it even though it is unused here, otherwise ACT
    raises TypeError: HighFiveOut() takes no arguments (1 given).
    """
    _harden_console_output()
    _prepare_environment()
    xaml_path = os.path.join(PROJECT_DIR, "AnsysReportGenerator_WPF.xaml")
    app = ReportGeneratorApp(xaml_path)
    app.window.ShowDialog()
