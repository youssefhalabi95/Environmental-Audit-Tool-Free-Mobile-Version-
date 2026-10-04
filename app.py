from flask import Flask, render_template, request, send_file, redirect, url_for
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime

app = Flask(__name__)


# ---------------------------------------------------------
# PREDEFINED AUDIT CHECKLIST
# ---------------------------------------------------------

CHECKLIST = {

    "Waste Management": [
        "Are waste streams properly identified and segregated?",
        "Are waste containers clearly labelled according to the waste type?",
        "Are waste storage areas appropriately designated and maintained?",
        "Are hazardous and non-hazardous wastes stored separately?",
        "Are waste storage areas protected from rain, runoff and unauthorized access?",
        "Are waste collection and disposal activities properly documented?",
        "Are wastes transferred only to authorized waste contractors or facilities?",
        "Are waste manifests, transfer notes or disposal records maintained?",
        "Are waste containers in good condition and free from significant damage or leakage?",
        "Are opportunities for waste reduction, reuse and recycling being implemented?"
    ],

    "Hazardous Materials": [
        "Are hazardous materials stored in designated areas?",
        "Are chemical containers properly labelled?",
        "Are Safety Data Sheets (SDS) available for applicable chemicals?",
        "Are incompatible chemicals stored separately?",
        "Are secondary containment measures provided where required?",
        "Are chemical storage areas protected from unauthorized access?",
        "Are spill response materials available near chemical storage areas?",
        "Are chemical containers inspected regularly for leakage or damage?",
        "Are expired or unused chemicals properly managed and disposed of?",
        "Are employees handling hazardous materials appropriately trained?"
    ],

    "Spill Prevention & Response": [
        "Are spill prevention measures implemented for fuels, oils and chemicals?",
        "Are spill kits available at relevant locations?",
        "Are spill kits adequately stocked and accessible?",
        "Are employees aware of the spill response procedure?",
        "Are significant spills and environmental incidents documented?",
        "Are contaminated materials generated from spills properly managed?",
        "Are drainage systems protected from potential chemical or oil releases?",
        "Are fuel and chemical transfer activities carried out in controlled areas?"
    ],

    "Air Quality & Dust": [
        "Are appropriate measures implemented to control dust generation?",
        "Are roads and exposed surfaces maintained to minimize dust?",
        "Are dusty materials appropriately covered during transportation?",
        "Are material stockpiles appropriately managed?",
        "Are vehicle speeds controlled in areas where dust generation is significant?",
        "Are air emission sources identified and appropriately controlled?",
        "Are required air quality or emission monitoring activities being conducted?",
        "Are monitoring results reviewed and corrective actions implemented where necessary?"
    ],

    "Noise": [
        "Are significant noise sources identified?",
        "Are noise-generating activities appropriately controlled?",
        "Are noisy activities restricted to appropriate working hours where required?",
        "Are equipment and machinery maintained to minimize excessive noise?",
        "Are noise monitoring activities conducted where required?",
        "Are noise complaints recorded and investigated?",
        "Are appropriate mitigation measures implemented where monitoring indicates elevated noise?"
    ],

    "Water & Wastewater": [
        "Are water consumption sources identified and monitored?",
        "Are wastewater discharge points identified?",
        "Are wastewater discharges managed in accordance with applicable requirements?",
        "Are wastewater treatment systems properly operated and maintained?",
        "Are wastewater monitoring results documented?",
        "Are stormwater drainage systems maintained and free from obstruction?",
        "Are potentially contaminated runoff pathways appropriately controlled?",
        "Are leaks or unauthorized discharges promptly investigated and addressed?"
    ],

    "Soil & Groundwater": [
        "Are activities with potential soil contamination risks identified?",
        "Are fuel and chemical storage areas appropriately protected?",
        "Are spills and leaks promptly cleaned and investigated?",
        "Are contaminated soils properly identified and managed?",
        "Are excavation and soil handling activities appropriately controlled?",
        "Are areas with historical contamination appropriately monitored or managed?"
    ],

    "Biodiversity": [
        "Have sensitive ecological receptors in the project area been identified?",
        "Are activities restricted in environmentally sensitive areas where required?",
        "Are vegetation clearance activities controlled and authorized?",
        "Are measures implemented to minimize impacts on wildlife?",
        "Are biodiversity-related monitoring requirements being implemented?",
        "Are ecological incidents or wildlife interactions documented?"
    ],

    "Environmental Permits & Legal Compliance": [
        "Are all required environmental permits and approvals available?",
        "Are permits and approvals valid and up to date?",
        "Are permit conditions being monitored for compliance?",
        "Are environmental regulatory reporting requirements being fulfilled?",
        "Are required environmental records maintained?",
        "Are applicable environmental laws and regulations identified?",
        "Are changes to applicable environmental requirements monitored?",
        "Are regulatory inspections and correspondence properly documented?"
    ],

    "Environmental Management System": [
        "Is there a documented environmental policy?",
        "Are environmental responsibilities clearly assigned?",
        "Are environmental objectives and targets established where appropriate?",
        "Are environmental procedures available to relevant personnel?",
        "Are environmental training records maintained?",
        "Are environmental incidents documented and investigated?",
        "Are corrective and preventive actions tracked to completion?",
        "Are environmental records properly maintained and controlled?",
        "Are environmental performance results periodically reviewed?"
    ],

    "Emergency Preparedness": [
        "Are environmental emergency scenarios identified?",
        "Is an environmental emergency response procedure available?",
        "Are emergency contact details readily available?",
        "Are spill response materials accessible?",
        "Are relevant employees trained in environmental emergency response?",
        "Are emergency drills conducted where required?",
        "Are lessons learned from environmental incidents incorporated into procedures?"
    ]
}


# ---------------------------------------------------------
# LANDING PAGE
# ---------------------------------------------------------

@app.route("/")
def landing():
    return render_template("landing.html")


# ---------------------------------------------------------
# AUDIT TOOL
# ---------------------------------------------------------

@app.route("/tool")
def tool():
    categories = list(CHECKLIST.keys())
    return render_template(
        "index.html",
        categories=categories,
        checklist=CHECKLIST
    )


# ---------------------------------------------------------
# EXCEL EXPORT
# ---------------------------------------------------------

@app.route("/export", methods=["POST"])
def export():

    project_name = request.form.get("project_name", "")
    site_name = request.form.get("site_name", "")
    auditor_name = request.form.get("auditor_name", "")
    audit_date = request.form.get("audit_date", "")
    
    wb = Workbook()

    # -----------------------------------------------------
    # Summary sheet
    # -----------------------------------------------------

    ws = wb.active
    ws.title = "Audit Summary"

    ws["A1"] = "Environmental Audit Checklist"
    ws["A3"] = "Project"
    ws["B3"] = project_name

    ws["A4"] = "Site"
    ws["B4"] = site_name

    ws["A5"] = "Auditor"
    ws["B5"] = auditor_name

    ws["A6"] = "Audit Date"
    ws["B6"] = audit_date

    # Count statuses
    statuses = request.form.getlist("status")
    
    compliant = statuses.count("Compliant")
    partial = statuses.count("Partially Compliant")
    non_compliant = statuses.count("Non-Compliant")
    na = statuses.count("N/A")

    ws["A8"] = "Summary"
    ws["A9"] = "Compliant"
    ws["B9"] = compliant

    ws["A10"] = "Partially Compliant"
    ws["B10"] = partial

    ws["A11"] = "Non-Compliant"
    ws["B11"] = non_compliant

    ws["A12"] = "N/A"
    ws["B12"] = na

    applicable = compliant + partial + non_compliant

    if applicable > 0:
        compliance_percentage = round(
            ((compliant + (partial * 0.5)) / applicable) * 100,
            1
        )
    else:
        compliance_percentage = 0

    ws["A14"] = "Compliance Score"
    ws["B14"] = f"{compliance_percentage}%"

    # -----------------------------------------------------
    # Detailed sheet
    # -----------------------------------------------------

    detail = wb.create_sheet("Audit Findings")

    headers = [
        "No.",
        "Category",
        "Audit Question",
        "Status",
        "Finding / Observation",
        "Corrective Action",
        "Responsible Person",
        "Target Date"
    ]

    for col, header in enumerate(headers, start=1):
        detail.cell(row=1, column=col).value = header

    questions = request.form.getlist("question")
    categories = request.form.getlist("category")
    findings = request.form.getlist("finding")
    actions = request.form.getlist("action")
    responsible = request.form.getlist("responsible")
    target_dates = request.form.getlist("target_date")

    for i in range(len(questions)):

        row = i + 2

        detail.cell(row=row, column=1).value = i + 1
        detail.cell(row=row, column=2).value = categories[i]
        detail.cell(row=row, column=3).value = questions[i]
        detail.cell(row=row, column=4).value = statuses[i]
        detail.cell(row=row, column=5).value = findings[i]
        detail.cell(row=row, column=6).value = actions[i]
        detail.cell(row=row, column=7).value = responsible[i]
        detail.cell(row=row, column=8).value = target_dates[i]

    # -----------------------------------------------------
    # Formatting
    # -----------------------------------------------------

    for sheet in wb.worksheets:

        for column in sheet.columns:

            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:

                if cell.value:
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            sheet.column_dimensions[column_letter].width = min(
                max_length + 2,
                50
            )

        sheet.freeze_panes = "A2"

    # -----------------------------------------------------
    # Create downloadable Excel file
    # -----------------------------------------------------

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    filename = "Environmental_Audit_Checklist.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ---------------------------------------------------------
# ABOUT
# ---------------------------------------------------------

@app.route("/about")
def about():
    return render_template("about.html")


# ---------------------------------------------------------
# PRIVACY POLICY
# ---------------------------------------------------------

@app.route("/privacy")
def privacy():
    return render_template("privacy.html")


# ---------------------------------------------------------
# RUN APPLICATION
# ---------------------------------------------------------

if __name__ == "__main__":
    app.run(debug=True)