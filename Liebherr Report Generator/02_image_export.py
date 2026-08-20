# 02_image_export.py: image export - capturing the Mechanical graphics view for geometry/mesh/BC/results, and rebuilding a chart from a CSV for Solution Information trackers. Depends on 00_constants.py (must be executed before this file).

import csv

import clr
clr.AddReference("System.Drawing")

from System.Drawing import (Bitmap, Color, Font, FontFamily, FontStyle, Graphics, Pen, PointF,
                             RectangleF, SolidBrush, StringAlignment, StringFormat)
from System.Drawing.Drawing2D import SmoothingMode
from System.Drawing.Imaging import ImageFormat

CHART_COLORS = [Color.IndianRed, Color.SteelBlue, Color.SeaGreen, Color.DarkOrange, Color.MediumPurple]


def _parse_float(text):
    """
    Does: converts a Mechanical cell text into a float, comma or dot decimal.
    Depends on: nothing (pure string processing).
    Returns: float, or None if not convertible.
    """
    # Mechanical runs in a French locale and returns values like "1,234E-05" in the Tabular Data pane.
    if text is None:
        return None
    try:
        return float(text.strip().replace(",", "."))
    except ValueError:
        return None


def export_current_view_image(image_name):
    """
    Does: exports the current graphics view to PNG (Ansys logo hidden) with the report's default settings.
    Depends on: ExtAPI.Graphics.ExportImage/ViewOptions (Ansys API), get_unique_file_path (00_constants.py).
    Returns: str, the path of the written PNG file.
    """
    # ShowLogo=False forced here (common pass-through point for all image exports) rather than
    # in each high-level export function: guarantees that no report image shows
    # the Ansys logo, without having to remember it in every new caller.
    ExtAPI.Graphics.ViewOptions.ShowLogo = False

    settings = Ansys.Mechanical.Graphics.GraphicsImageExportSettings()
    settings.CurrentGraphicsDisplay = False
    settings.Background = GraphicsBackgroundType.White
    settings.Width = DEFAULT_IMAGE_WIDTH
    settings.Height = DEFAULT_IMAGE_HEIGHT

    image_path = get_unique_file_path(IMAGE_EXPORT_FOLDER, image_name, ".png")
    ExtAPI.Graphics.ExportImage(image_path, GraphicsImageExportFormat.PNG, settings)
    return image_path


def set_material_display():
    """
    Does: switches the graphics view to coloring by material, mesh hidden.
    Depends on: ExtAPI.Graphics.ViewOptions (Ansys API).
    Returns: nothing (side effect on the display state).
    """
    ExtAPI.Graphics.ViewOptions.ModelColoring = ModelColoring.ByMaterial
    ExtAPI.Graphics.ViewOptions.ShowMesh = False


def export_geometry_image():
    """
    Does: exports the geometry view, colored by material.
    Depends on: set_material_display, ExtAPI.DataModel.Project.Model.Geometry, export_current_view_image.
    Returns: str, the path of the PNG.
    """
    set_material_display()
    ExtAPI.DataModel.Project.Model.Geometry.Activate()
    return export_current_view_image("geometry")


def export_mesh_image():
    """
    Does: exports the mesh view.
    Depends on: set_material_display, ExtAPI.DataModel.Project.Model.Mesh, export_current_view_image.
    Returns: str, the path of the PNG.
    """
    set_material_display()
    ExtAPI.DataModel.Project.Model.Mesh.Activate()
    return export_current_view_image("mesh")


def export_analysis_overview_image(analysis=None):
    """
    Does: exports the annotated overview view (BC A, B, C...) shown when selecting the analysis root in the tree.
    Depends on: export_current_view_image; analysis or, if None, ExtAPI.DataModel.Project.Model.Analyses[0].
    Returns: str, the path of the PNG.
    """
    # Stays generic regardless of analysis type by activating the given analysis instead of a hardcoded name.
    analysis = analysis or ExtAPI.DataModel.Project.Model.Analyses[0]
    analysis.Activate()
    return export_current_view_image(analysis.Name)


def export_object_image(obj, image_name):
    """
    Does: activates an object and exports its image via a "Figure" snapshot (more reliable than a direct export under repeated use).
    Depends on: obj.AddFigure() if available, otherwise falls back to Activate(); export_current_view_image.
    Returns: str, the path of the PNG.
    """
    # No SetFit() here: the camera framing (view chosen via apply_view_if_exists, or the
    # current manual position) is left as-is, at the user's responsibility - a
    # SetFit() would silently override any custom view right before the capture.
    add_figure = getattr(obj, "AddFigure", None)
    if add_figure is not None:
        try:
            figure = add_figure()
            figure.Activate()
            image_path = export_current_view_image(image_name)
            obj.Activate()  # restores the tree state for subsequent calls (e.g.: tabular data extraction)
            return image_path
        except Exception as e:
            print "AddFigure() failed for {} ({}): using direct export instead.".format(obj.Name, str(e))

    obj.Activate()
    return export_current_view_image(image_name)


def export_solution_image(result):
    """
    Does: exports the view of a given solution result.
    Depends on: export_object_image.
    Returns: str, the path of the PNG.
    """
    return export_object_image(result, result.Name)


def _read_chart_data(csv_path):
    """
    Does: reads a CSV (';' delimiter) and separates the optional header row from the numeric data rows.
    Depends on: the csv module, _parse_float.
    Returns: tuple (headers, rows) - headers is a list of column names (None if absent), rows a list of float rows.
    """
    # The 1st row is treated as a header if it isn't fully convertible to numbers.
    raw_rows = []
    with open(csv_path, "rb") as f:
        reader = csv.reader(f, delimiter=';')
        for row in reader:
            if row:
                raw_rows.append(row)

    if not raw_rows:
        return None, []

    headers = None
    data_rows = raw_rows
    first_row_values = [_parse_float(cell) for cell in raw_rows[0]]
    if not all(v is not None for v in first_row_values):
        headers = raw_rows[0]
        data_rows = raw_rows[1:]

    rows = []
    for row in data_rows:
        values = [_parse_float(cell) for cell in row]
        if values and all(v is not None for v in values):
            rows.append(values)

    return headers, rows


def export_chart_image_from_csv(csv_path, image_name, chart_title=None, x_axis_label=None,
                                 y_axis_label=None, curve_color=None):
    """
    Does: builds a chart image (title, graduated axes, grid, curves + points) from a CSV exported by export_active_tabular_data.
    Depends on: _read_chart_data, System.Drawing (Bitmap/Graphics/Pen...), CHART_COLORS.
    Returns: str, the path of the generated PNG, or None if the CSV doesn't have enough usable numeric data.
    """
    # Used for objects that only display a 2D chart in Mechanical (Solution Information trackers): no 3D view to capture, so the chart is redrawn from the exported data.
    # 1st column = X axis (Time/Step) and each following column = a curve if the CSV has 2+ columns; otherwise the row number serves as the X axis.
    headers, rows = _read_chart_data(csv_path)
    if len(rows) < 2:
        print "Not enough numeric data to plot a chart: " + csv_path
        return None

    num_columns = min(len(r) for r in rows)
    rows = [r[:num_columns] for r in rows]

    if num_columns >= 2:
        x_values = [r[0] for r in rows]
        series = [[r[c] for r in rows] for c in range(1, num_columns)]
    else:
        x_values = [float(i + 1) for i in range(len(rows))]
        series = [[r[0] for r in rows]]

    if headers and len(headers) >= num_columns:
        x_label = headers[0] if num_columns >= 2 else "N"
        series_labels = [headers[c] for c in range(1, num_columns)] if num_columns >= 2 else [headers[0]]
    else:
        x_label = "X"
        series_labels = ["Series {}".format(i + 1) for i in range(len(series))]

    if x_axis_label:
        x_label = x_axis_label
    if y_axis_label and len(series) == 1:
        series_labels[0] = y_axis_label

    width, height = 900, 550
    plot_left, plot_top = 90, 50
    plot_right, plot_bottom = width - 30, height - 80

    x_min, x_max = min(x_values), max(x_values)
    y_all = [v for serie in series for v in serie]
    y_min, y_max = min(y_all), max(y_all)
    if x_max == x_min:
        x_max += 1.0
    if y_max == y_min:
        y_max += 1.0

    def to_pixel(x, y):
        # Does: converts a data point (x, y) into a pixel of the plotting area. Depends on: x_min/x_max/y_min/y_max, plot_left/top/right/bottom (enclosing scope). Returns: PointF.
        px = plot_left + (x - x_min) / (x_max - x_min) * (plot_right - plot_left)
        py = plot_bottom - (y - y_min) / (y_max - y_min) * (plot_bottom - plot_top)
        return PointF(px, py)

    bitmap = Bitmap(width, height)
    g = Graphics.FromImage(bitmap)
    try:
        g.Clear(Color.White)
        g.SmoothingMode = SmoothingMode.AntiAlias

        title_font = Font(FontFamily.GenericSansSerif, 14, FontStyle.Bold)
        axis_title_font = Font(FontFamily.GenericSansSerif, 10, FontStyle.Bold)
        label_font = Font(FontFamily.GenericSansSerif, 9)
        text_brush = SolidBrush(Color.Black)

        title_format = StringFormat()
        title_format.Alignment = StringAlignment.Center
        g.DrawString(chart_title if chart_title else image_name, title_font, text_brush,
                      RectangleF(0, 10, width, 30), title_format)

        # --- Fine grid (minor, ungraduated) ---
        minor_grid_pen = Pen(Color.FromArgb(240, 240, 240), 1)
        minor_divisions = 20
        for i in range(1, minor_divisions):
            t = float(i) / minor_divisions
            py = plot_bottom - t * (plot_bottom - plot_top)
            g.DrawLine(minor_grid_pen, plot_left, py, plot_right, py)
            px = plot_left + t * (plot_right - plot_left)
            g.DrawLine(minor_grid_pen, px, plot_top, px, plot_bottom)

        # --- Major grid + graduations (4 divisions on each axis) ---
        major_grid_pen = Pen(Color.Gainsboro, 1)
        divisions = 4
        for i in range(divisions + 1):
            t = float(i) / divisions

            y_val = y_min + t * (y_max - y_min)
            py = plot_bottom - t * (plot_bottom - plot_top)
            g.DrawLine(major_grid_pen, plot_left, py, plot_right, py)
            g.DrawString("{:.3g}".format(y_val), label_font, text_brush, plot_left - 75, py - 7)

            x_val = x_min + t * (x_max - x_min)
            px = plot_left + t * (plot_right - plot_left)
            g.DrawLine(major_grid_pen, px, plot_top, px, plot_bottom)
            g.DrawString("{:.3g}".format(x_val), label_font, text_brush, px - 20, plot_bottom + 8)

        # --- Axes ---
        axis_pen = Pen(Color.Black, 2)
        g.DrawLine(axis_pen, plot_left, plot_bottom, plot_right, plot_bottom)
        g.DrawLine(axis_pen, plot_left, plot_top, plot_left, plot_bottom)

        # --- Axis titles (CSV column names, or x_axis_label/y_axis_label overrides) ---
        x_title_format = StringFormat()
        x_title_format.Alignment = StringAlignment.Center
        g.DrawString(x_label, axis_title_font, text_brush,
                      RectangleF(plot_left, plot_bottom + 28, plot_right - plot_left, 20), x_title_format)

        if len(series) == 1:
            state = g.Save()
            g.TranslateTransform(22, (plot_top + plot_bottom) / 2.0)
            g.RotateTransform(-90)
            y_title_format = StringFormat()
            y_title_format.Alignment = StringAlignment.Center
            g.DrawString(series_labels[0], axis_title_font, text_brush, RectangleF(-100, -15, 200, 20), y_title_format)
            g.Restore(state)

        # --- Curves (+ marked points, + legend if multiple series) ---
        legend_y = plot_top
        for series_index in range(len(series)):
            serie = series[series_index]
            color = curve_color if curve_color is not None else CHART_COLORS[series_index % len(CHART_COLORS)]
            curve_pen = Pen(color, 3)
            marker_brush = SolidBrush(color)

            pixel_points = [to_pixel(x_values[i], serie[i]) for i in range(len(serie))]
            for i in range(len(pixel_points) - 1):
                g.DrawLine(curve_pen, pixel_points[i], pixel_points[i + 1])
            for p in pixel_points:
                g.FillEllipse(marker_brush, p.X - 4, p.Y - 4, 8, 8)

            if len(series) > 1:
                g.FillRectangle(marker_brush, plot_right - 140, legend_y, 12, 12)
                g.DrawString(series_labels[series_index], label_font, text_brush,
                              plot_right - 122, legend_y - 2)
                legend_y += 18

        image_path = get_unique_file_path(IMAGE_EXPORT_FOLDER, image_name, ".png")
        bitmap.Save(image_path, ImageFormat.Png)
    finally:
        g.Dispose()
        bitmap.Dispose()

    return image_path
