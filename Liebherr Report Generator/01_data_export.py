# 01_data_export.py: data extraction - everything that reads the model / the Tabular Data pane and writes CSV files. Depends on 00_constants.py (must be executed before this file).

import csv
import os
import materials


def export_active_tabular_data(directory, filename, start_col=3):
    """
    Does: exports the content of the currently displayed Tabular Data pane to a CSV file.
    Depends on: ExtAPI.UserInterface (active Mechanical pane); the caller must have called Activate() on the object before this call.
    Returns: str, the full path of the written CSV.
    """
    # start_col skips the first columns of the pane (line number / step), not relevant for the export.
    pane = ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
    control = pane.ControlUnknown
    num_columns = control.ColumnsCount + 1
    num_rows = control.RowsCount + 1

    rows = []
    for row in range(1, num_rows):
        line = [clean_cell_text(control.cell(row, col).Text) for col in range(start_col, num_columns)]
        if any(cell != "" for cell in line):
            rows.append(line)

    filepath = os.path.join(directory, filename)
    with open(filepath, 'wb') as f:  # 'wb' to avoid double line breaks on Windows
        writer = csv.writer(f, delimiter=';')
        for line in rows:
            writer.writerow([to_csv_cell(cell) for cell in line])

    print "CSV exported: " + filepath
    return filepath


def export_bc_tabular_data(directory, bc):
    """
    Does: exports the tabular data of a given Boundary Condition.
    Depends on: export_active_tabular_data (after Activate() on bc).
    Returns: str, the path of the CSV.
    """
    bc.Activate()
    return export_active_tabular_data(directory, "{}.csv".format(bc.Name), start_col=3)


def export_bp_tabular_data(directory, bp):
    """
    Does: exports the tabular data of a given Bolt Pretension.
    Depends on: export_active_tabular_data (after Activate() on bp).
    Returns: str, the path of the CSV.
    """
    bp.Activate()
    return export_active_tabular_data(directory, "{}.csv".format(bp.Name), start_col=3)


def export_result_tabular_data(directory, obj):
    """
    Does: exports the tabular data of any result object (solution result, Contact/Bolt Tool child, Solution Information tracker...).
    Depends on: export_active_tabular_data; obj must support .Activate() and .Name.
    Returns: str, the path of the CSV.
    """
    # All these object types share the same Tabular Data pane layout (column 1 ignored, data starting from column 2).
    obj.Activate()
    return export_active_tabular_data(directory, "{}.csv".format(obj.Name), start_col=2)


def export_contacts_summary_csv(directory, contact_list=None):
    """
    Does: exports a summary table (type, friction, stiffness, tolerances, interface treatment) for each Contact Region in the model.
    Depends on: ExtAPI.DataModel (if contact_list is None) and _get_prop to read properties safely.
    Returns: str, the path of the CSV.
    """
    filepath = os.path.join(directory, "contact_info_export.csv")
    if contact_list is None:
        # None = historical behavior: all Contact Regions in the model (see create_contact_summary_slide in 04_slides.py).
        contact_list = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactRegion)

    with open(filepath, 'wb') as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
            "Name", "Contact Type", "Friction Coefficient", "Normal Stiffness Factor",
            "Penetration Tolerance", "Penetration Tolerance Value",
            "Elastic Slip Tolerance", "Elastic Slip Tolerance Value",
            "Interface Treatment", "Offset",
        ])

        for contact in contact_list:
            try:
                friction = ""
                if contact.ContactType == ContactType.Frictional:
                    friction = contact.FrictionCoefficient

                penetration_tolerance = _get_prop(contact, "PenetrationTolerance")
                penetration_tolerance_value = ""
                if penetration_tolerance is not None and "Value" in str(penetration_tolerance):
                    penetration_tolerance_value = _get_prop(contact, "PenetrationToleranceValue")

                elastic_slip_tolerance = _get_prop(contact, "ElasticSlipTolerance")
                elastic_slip_tolerance_value = ""
                if elastic_slip_tolerance is not None and "Value" in str(elastic_slip_tolerance):
                    elastic_slip_tolerance_value = _get_prop(contact, "ElasticSlipToleranceValue")

                interface_treatment = _get_prop(contact, "InterfaceTreatment")
                offset_value = ""
                if interface_treatment is not None and "Offset" in str(interface_treatment):
                    offset_value = _get_prop(contact, "Offset")

                writer.writerow([
                    to_csv_cell(contact.Name), to_csv_cell(contact.ContactType),
                    to_csv_cell(friction), to_csv_cell(contact.NormalStiffnessFactor),
                    to_csv_cell(penetration_tolerance), to_csv_cell(penetration_tolerance_value),
                    to_csv_cell(elastic_slip_tolerance), to_csv_cell(elastic_slip_tolerance_value),
                    to_csv_cell(interface_treatment), to_csv_cell(offset_value),
                ])
            except Exception as e:
                print "Error on contact {}: {}".format(contact.Name, str(e))

    print "CSV export complete: " + filepath
    return filepath


def export_mesh_report_csv(directory):
    """
    Does: exports a complete report of mesh parameters and statistics (defaults, sizing, quality, inflation, advanced, statistics).
    Depends on: Model.Mesh and _get_prop to read properties safely, _format_element_size for ElementSize.
    Returns: str, the path of the CSV.
    """
    mesh = Model.Mesh
    rows = []

    rows.append(["Defaults", "PhysicsPreference", mesh.PhysicsPreference])
    rows.append(["Defaults", "ElementOrder", mesh.ElementOrder])
    rows.append(["Defaults", "ElementSize", _format_element_size(mesh)])

    for p in ["UseAdaptiveSizing", "Resolution", "MeshDefeaturing", "DefeatureSize",
              "TransitionOption", "SpanAngleCenter", "CurvatureNormalAngle",
              "MinSize", "MaxSize", "GrowthRate"]:
        rows.append(["Sizing", p, _get_prop(mesh, p)])

    for p in ["MeshMetric", "ErrorLimits", "TargetQuality", "Smoothing"]:
        rows.append(["Quality", p, _get_prop(mesh, p)])

    for p in ["UseAutomaticInflation", "InflationOption", "NumberOfLayers", "InflationGrowthRate"]:
        rows.append(["Inflation", p, _get_prop(mesh, p)])

    for p in ["NumberOfCPUsForParallelPartMeshing", "StraightSidedElements", "RigidBodyBehavior"]:
        rows.append(["Advanced", p, _get_prop(mesh, p)])

    rows.append(["Statistics", "Nodes", mesh.Nodes])
    rows.append(["Statistics", "Elements", mesh.Elements])

    try:
        metric_data = mesh.MeshMetricValues
        if metric_data and len(metric_data) > 0:
            rows.append(["Statistics", "MeshMetricValueMin", min(metric_data)])
            rows.append(["Statistics", "MeshMetricValueMax", max(metric_data)])
            rows.append(["Statistics", "MeshMetricValueAvg", sum(metric_data) / len(metric_data)])
    except Exception:
        rows.append(["Statistics", "MeshMetricValues", "Not available"])

    filepath = os.path.join(directory, "mesh_report.csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Section", "Property", "Value"])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "CSV export complete: " + filepath
    return filepath


def _get_prop(obj, prop_name):
    """
    Does: retrieves a property of a Mechanical object safely.
    Depends on: Python's native getattr.
    Returns: the property value, or None if absent / if accessing it raises an exception.
    """
    try:
        return getattr(obj, prop_name)
    except Exception:
        return None


def _format_element_size(mesh):
    """
    Does: reads ElementSize and distinguishes "Default" sizing (no value entered by the user,
    Quantity is zero) from an actually defined value.
    Depends on: _get_prop, mesh.ElementSize (.NET Quantity, exposes .Value) - when Element Size is
    left on "Default" in Mechanical, the actual size is computed dynamically at meshing time and
    is never written to this property, which then stays at 0.
    Returns: str "Default" if ElementSize is 0, otherwise the raw Quantity (handled normally by to_csv_cell).
    """
    element_size = _get_prop(mesh, "ElementSize")
    try:
        if element_size is not None and float(element_size.Value) == 0.0:
            return "Default"
    except Exception:
        pass
    return element_size


def export_materials_csv(directory):
    """
    Does: exports one row per distinct material used by the model's bodies (modulus, density, Poisson's ratio, thermal expansion, conductivity, specific heat).
    Depends on: ExtAPI.DataModel, materials module (Ansys API), _material_property_values / _material_property_units.
    Returns: str, the path of the CSV.
    """
    bodies = ExtAPI.DataModel.Project.Model.GetChildren(DataModelObjectCategory.Body, True)

    seen_materials = []
    for body in bodies:
        if body.Material not in seen_materials:
            seen_materials.append(body.Material)

    filepath = os.path.join(directory, "materials_export.csv")

    with open(filepath, 'wb') as f:
        first_material = bodies[0].GetGeoBody().Material
        units = _material_property_units(first_material)

        header = (
            "Material;Young's Modulus [{}];Density [{}];Poisson's Ratio [-];"
            "Thermal Expansion [{}];Thermal Conductivity [{}];Specific Heat [{}]\n"
        ).format(units["Young's Modulus"], units["Density"], units["Coefficient of Thermal Expansion"],
                 units["Thermal Conductivity"], units["Specific Heat"])
        f.write(to_csv_cell(header))

        for mat in seen_materials:
            body_for_mat = None
            for body in bodies:
                if body.Material == mat:
                    body_for_mat = body
                    break

            if body_for_mat is None:
                continue

            material = body_for_mat.GetGeoBody().Material
            values = _material_property_values(material)

            line = "{};{};{};{};{};{};{}\n".format(
                mat,
                values["Young's Modulus"],
                values["Density"],
                values["Poisson's Ratio"],
                values["Coefficient of Thermal Expansion"],
                values["Thermal Conductivity"],
                values["Specific Heat"],
            )
            f.write(to_csv_cell(line))

    print "CSV export complete: " + filepath
    return filepath


def _convert_young_modulus_to_gpa(prop_name, unit, value):
    """
    Does: converts Young's modulus from Pa to GPa (more readable in the materials table of the geometry slide).
    Depends on: nothing (simple numeric conversion); only applies if prop_name == "Young's Modulus" and
    the source unit is "Pa", so a value already retrieved in another unit is never re-converted
    (e.g.: project configured in MPa/psi).
    Returns: tuple (unit, value) converted if applicable, otherwise (unit, value) unchanged.
    """
    if prop_name == "Young's Modulus" and unit == "Pa":
        converted_value = (value / 1.0e9) if value is not None else None
        return "GPa", converted_value
    return unit, value


def _material_property_values(material):
    """
    Does: flattens the 5 material property groups used in the report into a single dict of values.
    Depends on: materials.GetMaterialPropertyByName (Ansys API), _convert_young_modulus_to_gpa.
    Returns: dict, property name -> value (first value only if the property is temperature-dependent; Young's Modulus converted to GPa if retrieved in Pa).
    """
    values = {}
    for group in ["Elasticity", "Density", "Coefficient of Thermal Expansion",
                  "Thermal Conductivity", "Specific Heat"]:
        for prop_name, prop_data in materials.GetMaterialPropertyByName(material, group).items():
            # prop_data = (unit, value) for a constant property, or (unit, value_T1, value_T2, ...) if temperature-dependent: only the 1st value is kept, to stay at one row per material.
            unit = prop_data[0]
            value = prop_data[1] if len(prop_data) > 1 else prop_data[0]
            _, value = _convert_young_modulus_to_gpa(prop_name, unit, value)
            values[prop_name] = value
    return values


def _material_property_units(material):
    """
    Does: flattens the 5 material property groups used in the report into a single dict of units.
    Depends on: materials.GetMaterialPropertyByName (Ansys API), _convert_young_modulus_to_gpa.
    Returns: dict, property name -> unit ("GPa" for Young's Modulus if retrieved in Pa).
    """
    units = {}
    for group in ["Elasticity", "Density", "Coefficient of Thermal Expansion",
                  "Thermal Conductivity", "Specific Heat"]:
        for prop_name, prop_data in materials.GetMaterialPropertyByName(material, group).items():
            unit, _ = _convert_young_modulus_to_gpa(prop_name, prop_data[0], None)
            units[prop_name] = unit
    return units


def export_analysis_settings_csv(directory, analysis):
    """
    Does: exports a table of the analysis's step parameters (Analysis Settings), transposed (one Loadcase per column, one property per row) to avoid the repetitiveness of a classic vertical table.
    Depends on: analysis.AnalysisSettings (Ansys API), get_unique_file_path/safe_file_name/to_csv_cell (00_constants.py), the csv module.
    Returns: str, the path of the generated CSV.
    """
    settings = analysis.AnalysisSettings

    try:
        num_steps = settings.NumberOfSteps
    except Exception:
        num_steps = 0

    end_times = []
    define_bys = []
    auto_steppings = []
    substep_counts = []

    for step in range(1, num_steps + 1):
        try:
            end_times.append(settings.GetStepEndTime(step))
        except Exception:
            end_times.append(None)
        try:
            define_bys.append(settings.GetDefineBy(step))
        except Exception:
            define_bys.append(None)
        try:
            auto_steppings.append(settings.GetAutomaticTimeStepping(step))
        except Exception:
            auto_steppings.append(None)
        try:
            substep_counts.append(settings.GetNumberOfSubSteps(step))
        except Exception:
            substep_counts.append(None)

    header = ["Property"] + ["Loadcase {}".format(step) for step in range(1, num_steps + 1)]
    rows = [
        ["End time"] + end_times,
        ["Define by"] + define_bys,
        ["Auto time stepping"] + auto_steppings,
        ["Substeps"] + substep_counts,
    ]

    filepath = get_unique_file_path(directory, "AnalysisSettings_" + safe_file_name(analysis.Name), ".csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([to_csv_cell(v) for v in header])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "CSV export complete: " + filepath
    return filepath


SOLUTION_INFO_PROPERTIES = [  # Solution property name -> displayed label; list (not a dict) to keep a stable display order
    ("ElapsedRunTime", "Elapsed run time"),
    ("MemoryUsed", "Memory used"),
    ("ResultFileSize", "Result file size"),
]


def export_solution_info_csv(directory, solution, analysis_name):
    """
    Does: exports a table of solution information (run time, memory used, result file size).
    Depends on: solution.PropertyByName (Ansys API), SOLUTION_INFO_PROPERTIES constant, get_unique_file_path/safe_file_name/to_csv_cell (00_constants.py), the csv module.
    Returns: str, the path of the generated CSV.
    """
    rows = []
    for prop_name, label in SOLUTION_INFO_PROPERTIES:
        try:
            rows.append([label, solution.PropertyByName(prop_name).StringValue])
        except Exception:
            pass

    # analysis_name (not solution.Name, which is generic like "Solution" across all analyses) for a
    # distinct file name per analysis in a multi-analysis project.
    filepath = get_unique_file_path(directory, "SolutionInfo_" + safe_file_name(analysis_name), ".csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Property", "Value"])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "CSV export complete: " + filepath
    return filepath
