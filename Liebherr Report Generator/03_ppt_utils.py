# 03_ppt_utils.py: PowerPoint report builder - owns the COM Interop session and exposes "add slide" methods based on the corporate template layouts. Depends on 00_constants.py (must be executed before this file).

import clr
import csv
import datetime
import shutil

clr.AddReference("Microsoft.Office.Interop.PowerPoint")
clr.AddReference("Office")

import Microsoft.Office.Interop.PowerPoint as PPT
import Microsoft.Office.Core as Office


def _build_working_copy_base_name():
    """
    Does: builds a base name (without extension), unique per day.
    Depends on: datetime.date.today().
    Returns: str, e.g. "automatic_report_generation_17072025".
    """
    today = datetime.date.today()
    return "automatic_report_generation_{:02d}{:02d}{:04d}".format(today.day, today.month, today.year)


def rename_time_header_to_loadcase(data):
    """
    Does: replaces, in the header row of a CSV table, any cell equal exactly to "Time [s]" with "Loadcase".
    Depends on: nothing (modifies data in place).
    Returns: nothing (side effect on data).
    """
    # More meaningful to an engineer than the raw column name returned by Mechanical's Tabular Data pane.
    if not data:
        return
    header = data[0]
    for i in range(len(header)):
        if header[i] == u"Time [s]":
            header[i] = u"Loadcase"


class PPTReportBuilder(object):
    """
    Wraps a single PowerPoint Interop session opened on a working copy of the corporate template. All add_..._slide methods add slides to the SAME presentation instead of reopening PowerPoint for every slide.
    """

    def __init__(self, template_path):
        """
        Does: creates a working copy of the template and opens a PowerPoint COM session on it.
        Depends on: shutil.copyfile, get_unique_file_path (00_constants.py), Microsoft.Office.Interop.PowerPoint.
        Returns: nothing (initializes self.app / self.presentation / self.working_copy_path).
        """
        # The original template is NEVER opened directly: if the user hits Ctrl+S in PowerPoint during generation, this copy gets overwritten, never the original.
        self.working_copy_path = get_unique_file_path(
            REPORT_OUTPUT_FOLDER, _build_working_copy_base_name(), ".pptx")
        shutil.copyfile(template_path, self.working_copy_path)
        print "Template working copy opened: " + self.working_copy_path

        self.app = PPT.ApplicationClass()
        # self.app.Visible = True is necessary: a session left invisible has proven unstable on reports with many slides (COMException "Presentation.SlideMaster: Object does not exist" partway through, after which no more slides can be added/saved). The window closes normally at the end of generation (see close()).
        self.app.Visible = True
        self.presentation = self.app.Presentations.Open(self.working_copy_path, WithWindow=True)

    def _add_slide(self, layout_index):
        """
        Does: adds a blank slide at the end of the presentation, using the given custom layout.
        Depends on: self.presentation.SlideMaster.CustomLayouts.
        Returns: PPT.Slide, the created slide.
        """
        layout = self.presentation.SlideMaster.CustomLayouts[layout_index]
        return self.presentation.Slides.AddSlide(self.presentation.Slides.Count + 1, layout)

    def add_image_table_slide(self, title, subtitle, img_path=None, csv_path=None, comment=" "):
        """
        Does: adds an "image + table" slide (title, subtitle, one image, one table, one comment).
        Depends on: self._add_slide (LAYOUT_IMAGE_TABLE), self.add_csv_table.
        Returns: PPT.Slide, the created slide.
        """
        slide = self._add_slide(LAYOUT_IMAGE_TABLE)

        # Text assigned AFTER the slide is created: doing it directly on the layout would modify the whole master template.
        slide.Shapes[8].TextFrame.TextRange.Text = comment
        slide.Shapes[2].TextFrame.TextRange.Text = title
        slide.Shapes[4].TextFrame.TextRange.Text = subtitle

        if img_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[3]
            try:
                slide.Shapes.AddPicture(img_path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                         coord.Left, coord.Top, coord.Width, coord.Height)
            except Exception as e:
                print "Unable to insert image ({}): {}".format(img_path, str(e))

        if csv_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[1]
            try:
                self.add_csv_table(slide, csv_path, coord.Left, coord.Top, coord.Width)
            except Exception as e:
                print "Unable to insert table ({}): {}".format(csv_path, str(e))

        return slide

    def add_table_slide(self, title, subtitle, csv_path):
        """
        Does: adds a "table only" slide (title, subtitle, table).
        Depends on: self._add_slide (LAYOUT_TABLE_ONLY), self.add_csv_table.
        Returns: PPT.Slide, the created slide.
        """
        slide = self._add_slide(LAYOUT_TABLE_ONLY)
        slide.Shapes[1].TextFrame.TextRange.Text = title
        slide.Shapes[3].TextFrame.TextRange.Text = subtitle

        coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_TABLE_ONLY].Shapes[2]
        try:
            self.add_csv_table(slide, csv_path, coord.Left, coord.Top, coord.Width)
        except Exception as e:
            print "Unable to insert table ({}): {}".format(csv_path, str(e))
        return slide

    def add_analysis_context_slide(self, title, subtitle, img_path, settings_csv_path, solution_csv_path):
        """
        Does: adds the "Analysis Parameters" context slide (overview image + 2 tables: step settings, solution info).
        Depends on: self._add_slide (LAYOUT_IMAGE_TABLE), self.add_csv_table.
        Returns: PPT.Slide, the created slide.
        """
        # Reuses the LAYOUT_IMAGE_TABLE layout: its "table" placeholder (shape 1) receives the steps
        # table (same position as on the other image+table slides), and its "comment" placeholder
        # (shape 8, smaller) receives the second table (solution info, 3 rows max).
        slide = self._add_slide(LAYOUT_IMAGE_TABLE)

        slide.Shapes[2].TextFrame.TextRange.Text = title
        slide.Shapes[4].TextFrame.TextRange.Text = subtitle

        if img_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[3]
            try:
                slide.Shapes.AddPicture(img_path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                         coord.Left, coord.Top, coord.Width, coord.Height)
            except Exception as e:
                print "Unable to insert image ({}): {}".format(img_path, str(e))

        if settings_csv_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[1]
            try:
                self.add_csv_table(slide, settings_csv_path, coord.Left, coord.Top, coord.Width)
            except Exception as e:
                print "Unable to insert table ({}): {}".format(settings_csv_path, str(e))

        if solution_csv_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[8]
            try:
                self.add_csv_table(slide, solution_csv_path, coord.Left, coord.Top, coord.Width)
            except Exception as e:
                print "Unable to insert table ({}): {}".format(solution_csv_path, str(e))

        return slide

    def add_csv_table(self, slide, csv_path, left, top, width):
        """
        Does: reads a CSV (delimiter ';') and inserts it as a formatted table on the slide (bold/gray header, thin borders, centered text).
        Depends on: the csv module, rename_time_header_to_loadcase, MAX_TABLE_ROWS/MAX_TABLE_COLUMNS (00_constants.py).
        Returns: nothing (modifies slide; does nothing if the CSV is empty or exceeds the size limits).
        """
        data = []
        # Binary read + explicit UTF-8 decoding: CSVs are written in UTF-8 (units with special characters like degree/micro/superscripts); a text read without explicit encoding raises a DecoderFallbackException.
        with open(csv_path, "rb") as f:
            reader = csv.reader(f, delimiter=';')
            for row in reader:
                data.append([cell.decode("utf-8") for cell in row])

        if not data:
            print "Empty CSV, no table inserted: " + csv_path
            return

        rename_time_header_to_loadcase(data)

        rows = len(data)
        cols = len(data[0])

        if rows > MAX_TABLE_ROWS or cols > MAX_TABLE_COLUMNS:
            print ("The table exceeds 50 rows / 50 columns ({} rows, {} columns): it will "
                   "therefore not be shown in PowerPoint but is available in csv format at "
                   "this location: {}").format(rows, cols, csv_path)
            return

        table = slide.Shapes.AddTable(rows, cols, left, top, width).Table

        # Borders set ONCE PER ROW (Rows(r).Cells.Borders accepts a range of cells), not per cell x per side: every COM round-trip is costly, and this was the slowest part of formatting (up to 45s for 8 rows before this optimization). Font/TextRange, however, still have to be set per cell (no range equivalent).
        for r in range(1, rows + 1):
            row_cells = table.Rows(r).Cells
            for border_index in range(1, 5):
                row_cells.Borders(border_index).ForeColor.RGB = 0x000000
                row_cells.Borders(border_index).Weight = 1

        for r in range(rows):
            for c in range(cols):
                cell = table.Cell(r + 1, c + 1)
                shape = cell.Shape
                text_frame = shape.TextFrame
                text_range = text_frame.TextRange
                text_range.Text = data[r][c]
                text_range.Font.Size = 7
                text_range.ParagraphFormat.Alignment = 2  # 2 = center
                text_frame.VerticalAnchor = 3  # 3 = middle
                # Top/bottom margin reduced to 3pt (left/right margins unchanged): it's this internal margin, not the font size, that prevents PowerPoint from shrinking row height below a certain threshold.
                text_frame.MarginTop = 3
                text_frame.MarginBottom = 3

                if r == 0:
                    text_range.Font.Bold = True
                    shape.Fill.ForeColor.RGB = 0x545454
                    shape.Fill.Solid()

        # Height forced to a deliberately too-small value: PowerPoint automatically brings it back to the real minimum needed to fit the text - the only way to tighten a table that's already been created (AddTable allocates by default a height well above what's needed for size-7 text, which can overflow the slide).
        for r in range(1, rows + 1):
            table.Rows(r).Height = 1

    def save_as(self, output_path):
        """
        Does: saves the presentation to a new file (the original template is never overwritten).
        Depends on: self.presentation.SaveAs.
        Returns: nothing (side effect: creates/overwrites output_path).
        """
        self.presentation.SaveAs(output_path)
        print "Report saved: " + output_path

    def close(self):
        """
        Does: saves the working copy, closes the presentation and quits the PowerPoint application.
        Depends on: self.presentation.Save/Close, self.app.Quit.
        Returns: nothing (side effect; called when generation fails, so a PowerPoint session is never left invisible in memory).
        """
        self.presentation.Save()  # Simple Save(), not SaveAs: the file already has its final name/path (working_copy_path)
        self.presentation.Close()
        self.app.Quit()

    def keep_open(self):
        """
        Does: does nothing to the presentation (no Save, no Close, no Quit) - leaves it as-is, open and displayed.
        Depends on: nothing.
        Returns: nothing (side effect: none; exists to make explicit, at the end of a successful generation, the choice to leave the report open without saving it).
        """
        pass
