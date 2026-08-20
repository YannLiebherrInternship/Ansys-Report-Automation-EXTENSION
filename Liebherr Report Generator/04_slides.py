# 04_slides.py: slide construction - each create_..._slide(report) extracts the data it needs (CSV + image) then adds a slide to the given PPTReportBuilder instance. Depends on 00_constants.py, 01_data_export.py, 02_image_export.py and 03_ppt_utils.py (must be executed before this file).


def create_geometry_slide(report):
    """
    Does: adds the geometry + materials context slide.
    Depends on: export_geometry_image, export_materials_csv, report.add_image_table_slide.
    Returns: nothing (side effect on report).
    """
    img_path = export_geometry_image()
    csv_path = export_materials_csv(CSV_EXPORT_FOLDER)
    report.add_image_table_slide("Geometry and materials details", "-- Geometry and materials --",
                                  img_path=img_path, csv_path=csv_path)


def create_mesh_slide(report):
    """
    Does: adds the mesh context slide.
    Depends on: export_mesh_image, export_mesh_report_csv, report.add_image_table_slide.
    Returns: nothing (side effect on report).
    """
    img_path = export_mesh_image()
    csv_path = export_mesh_report_csv(CSV_EXPORT_FOLDER)
    report.add_image_table_slide("Mesh and mesh details", "-- Mesh --",
                                  img_path=img_path, csv_path=csv_path)


def create_analysis_parameters_slide(report, analysis=None):
    """
    Does: adds the Analysis Parameters context slide (overview image + steps table + solution info table) for a given analysis.
    Depends on: export_analysis_overview_image, export_analysis_settings_csv, export_solution_info_csv, report.add_analysis_context_slide.
    Returns: nothing (side effect on report).
    """
    # analysis=None (original behavior): uses Analyses[0], kept for callers without an argument (e.g. obsolete code AnsysReportGenerator_GUI.py).
    analysis = analysis or ExtAPI.DataModel.Project.Model.Analyses[0]
    img_path = export_analysis_overview_image(analysis)
    settings_csv_path = export_analysis_settings_csv(CSV_EXPORT_FOLDER, analysis)
    solution_csv_path = export_solution_info_csv(CSV_EXPORT_FOLDER, analysis.Solution, analysis.Name)
    report.add_analysis_context_slide(analysis.Name, "Analysis Parameters",
                                       img_path, settings_csv_path, solution_csv_path)


def create_bc_slide(report):
    """
    Does: adds a slide for each Boundary Condition found in the model.
    Depends on: ExtAPI.DataModel, export_object_image, export_bc_tabular_data, report.add_image_table_slide.
    Returns: nothing (side effect on report; does nothing if no BC is found).
    """
    bc_list = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.GenericBoundaryCondition)
    if not bc_list:
        print "No Boundary Condition: slide skipped."
        return
    for bc in bc_list:
        img_path = export_object_image(bc, bc.Name)
        csv_path = export_bc_tabular_data(CSV_EXPORT_FOLDER, bc)
        report.add_image_table_slide(bc.Name, "-- Boundary Conditions --",
                                      img_path=img_path, csv_path=csv_path)


def create_bp_slide(report):
    """
    Does: adds a slide for each Bolt Pretension found in the model.
    Depends on: ExtAPI.DataModel, export_object_image, export_bp_tabular_data, report.add_image_table_slide.
    Returns: nothing (side effect on report; does nothing if no Bolt Pretension is found).
    """
    bp_list = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltPretension)
    if not bp_list:
        print "No Bolt Pretension: slide skipped."
        return
    for bp in bp_list:
        img_path = export_object_image(bp, bp.Name)
        csv_path = export_bp_tabular_data(CSV_EXPORT_FOLDER, bp)
        report.add_image_table_slide(bp.Name, "-- Bolt Pretension --",
                                      img_path=img_path, csv_path=csv_path)


def create_contact_summary_slide(report):
    """
    Does: adds the contacts summary slide (table only, no image).
    Depends on: export_contacts_summary_csv, report.add_table_slide.
    Returns: nothing (side effect on report).
    """
    csv_path = export_contacts_summary_csv(CSV_EXPORT_FOLDER)
    report.add_table_slide("Contacts summary", "-- Contact --", csv_path)


def create_tool_children_slides(report, category, subtitle, include_table=False):
    """
    Does: adds a slide for each CHILD of each tool object found for the given category (the results to export are the tool's children, not the tool itself).
    Depends on: ExtAPI.DataModel, export_object_image, export_result_tabular_data (if include_table), report.add_image_table_slide.
    Returns: nothing (side effect on report; skips tools with no object found or no children).
    """
    tools = ExtAPI.DataModel.GetObjectsByType(category)
    if not tools:
        print "No object found for: " + subtitle
        return

    for tool in tools:
        children = tool.Children
        if children is None or len(children) == 0:
            print "No child under " + tool.Name + ": skipped."
            continue
        for child in children:
            try:
                img_path = export_object_image(child, child.Name)
            except Exception as e:
                print "Unable to export image for {}: {}".format(child.Name, str(e))
                img_path = None

            csv_path = None
            if include_table:
                try:
                    csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, child)
                except Exception as e:
                    print "Unable to export CSV for {}: {}".format(child.Name, str(e))

            if img_path or csv_path:
                report.add_image_table_slide(child.Name, subtitle, img_path=img_path, csv_path=csv_path)
            else:
                print "No exportable data for " + child.Name + ": slide skipped."


def create_contact_tool_slide(report):
    """
    Does: adds a slide (with tabular data) for each child result of each Contact Tool found.
    Depends on: create_tool_children_slides.
    Returns: nothing (side effect on report).
    """
    create_tool_children_slides(report, DataModelObjectCategory.ContactTool, "-- Contact Tool --", include_table=True)


def create_bolt_tool_slide(report):
    """
    Does: adds a slide (with tabular data) for each child result of each Bolt Tool found.
    Depends on: create_tool_children_slides.
    Returns: nothing (side effect on report).
    """
    create_tool_children_slides(report, DataModelObjectCategory.BoltTool, "-- Bolt Tool --", include_table=True)


def get_scoped_contact_region_name(obj):
    """
    Does: retrieves the name of the Contact Region associated with an object (Scope > Contact Region field), when it has one.
    Depends on: obj.ContactRegion (Ansys API).
    Returns: str, the Contact Region name, or None if not available.
    """
    try:
        contact_region = obj.ContactRegion
        if contact_region is None:
            return None
        return contact_region.Name
    except Exception:
        return None


def create_solution_information_slide(report):
    """
    Does: adds a slide for each tracker child of "Solution Information" (e.g. Pressure, Max Normal Stiffness, Elastic Slip).
    Depends on: ExtAPI.DataModel, export_result_tabular_data, export_chart_image_from_csv, get_scoped_contact_region_name, report.add_image_table_slide.
    Returns: nothing (side effect on report; does nothing if no tracker is found).
    """
    # These objects only display a 2D chart, no 3D view: we export the tabular data then rebuild the chart as an image from the CSV rather than capturing the viewport.
    analysis = ExtAPI.DataModel.Project.Model.Analyses[0]
    solution_information = analysis.Solution.Children[0]
    children = solution_information.Children

    if children is None or len(children) == 0:
        print "No tracker under Solution Information: slide skipped."
        return

    for child in children:
        csv_path = None
        try:
            csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, child)
        except Exception as e:
            print "Unable to export CSV for {}: {}".format(child.Name, str(e))

        img_path = None
        if csv_path:
            try:
                img_path = export_chart_image_from_csv(csv_path, child.Name)
            except Exception as e:
                print "Unable to build chart for {}: {}".format(child.Name, str(e))

        if img_path or csv_path:
            contact_region_name = get_scoped_contact_region_name(child)
            title = child.Name
            if contact_region_name:
                # Title suffixed with the scoped Contact Region (Details > Scope), to know which contact these values relate to.
                title = "{} - {}".format(child.Name, contact_region_name)

            report.add_image_table_slide(title, "-- Solution Information --",
                                          img_path=img_path, csv_path=csv_path)
        else:
            print "No exportable data for " + child.Name + ": slide skipped."


def get_all_simple_results():
    """
    Does: returns all simple results under the Solution branch (all children except Solution Information and the Contact/Bolt Tool folders, already handled elsewhere).
    Depends on: ExtAPI.DataModel.
    Returns: list, the remaining result objects.
    """
    excluded_categories = [DataModelObjectCategory.ContactTool, DataModelObjectCategory.BoltTool]

    # Exclusion by identity in addition to category: ensures any object already handled by create_contact_tool_slide / create_bolt_tool_slide is skipped here, even if its category doesn't match exactly.
    already_handled = list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactTool))
    already_handled += list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltTool))

    analysis = ExtAPI.DataModel.Project.Model.Analyses[0]
    children = analysis.Solution.Children

    results = []
    for i in range(1, len(children)):  # index 0 = Solution Information, excluded from the loop
        child = children[i]
        if child.DataModelObjectCategory in excluded_categories:
            continue
        if child in already_handled:
            continue
        results.append(child)
    return results


def create_result_slide(report):
    """
    Does: adds a slide for each simple result found under Solution.
    Depends on: get_all_simple_results, export_solution_image, export_result_tabular_data, report.add_image_table_slide.
    Returns: nothing (side effect on report; does nothing if no simple result is found).
    """
    analysis = ExtAPI.DataModel.Project.Model.Analyses[0]
    results = get_all_simple_results()
    if not results:
        print "No simple result: slide skipped."
        return
    for result in results:
        img_path = export_solution_image(result)
        csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, result)
        report.add_image_table_slide(result.Name, analysis.Name, img_path=img_path, csv_path=csv_path)
