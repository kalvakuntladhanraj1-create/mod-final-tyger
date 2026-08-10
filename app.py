from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
from datetime import datetime

app = Flask(__name__)


# ---------------- HOME ----------------
@app.route("/")
def home():
    return render_template("sale.html")


# ---------------- GENERATE REPORT ----------------
@app.route("/generate_sale", methods=["POST"])
def generate_sale():

    # -------- BASIC --------
    context = {
        "DATE": request.form.get("DATE"),
        "APPLICATION_NO": request.form.get("APPLICATION_NO"),
        "APPLICANT_NAME": request.form.get("APPLICANT_NAME"),
        "APPLICANT_BOLD_OWNER": request.form.get("APPLICANT_BOLD_OWNER"),
        "LOAN_AMOUNT": request.form.get("LOAN_AMOUNT"),

        # -------- PROPERTY --------
        "DOOR_NO": request.form.get("DOOR_NO"),
        "PLOT_NO": request.form.get("PLOT_NO"),
        "ASSESSMENT_NO": request.form.get("ASSESSMENT_NO"),
        "Assessment_No": request.form.get("ASSESSMENT_NO"),  # IMPORTANT (your template uses both)
        "EXTENT_YARDS": request.form.get("EXTENT_YARDS"),
        "SURVEY_NO": request.form.get("SURVEY_NO"),
        "ADDRESS": request.form.get("ADDRESS"),
        "VILLAGE": request.form.get("VILLAGE"),
        "GRAM_PANCHAYAT": request.form.get("GRAM_PANCHAYAT"),
        "MANDAL": request.form.get("MANDAL"),
        "SRO": request.form.get("SRO"),
        "RO": request.form.get("RO"),
        "DISTRICT": request.form.get("DISTRICT"),

        # -------- BOUNDARIES --------
        "EAST_BOUNDARY": request.form.get("EAST_BOUNDARY"),
        "WEST_BOUNDARY": request.form.get("WEST_BOUNDARY"),
        "NORTH_BOUNDARY": request.form.get("NORTH_BOUNDARY"),
        "SOUTH_BOUNDARY": request.form.get("SOUTH_BOUNDARY"),

        # -------- MEASUREMENTS --------
        "E_W": request.form.get("E_W"),
        "N_S": request.form.get("N_S"),
        "E_W_FEET": request.form.get("E_W_FEET"),
        "N_S_FEET": request.form.get("N_S_FEET"),
        "EXTENT_FEET": request.form.get("EXTENT_FEET"),

        # -------- DEED --------
        "DEED_NO": request.form.get("DEED_NO"),
        "DEED_DATE": request.form.get("DEED_DATE"),

        # -------- EXTRA --------
        "POSSESSION_DATE": request.form.get("POSSESSION_DATE"),
        "FROM_YEARS": request.form.get("FROM_YEARS"),
        "POSSESSION_NAME": request.form.get("POSSESSION_NAME"),

        "H_T_DATE": request.form.get("H_T_DATE"),
        "HOUSE_TAX_RECIPT_NO": request.form.get("HOUSE_TAX_RECIPT_NO"),
        "FINANCIAL_YEARS": request.form.get("FINANCIAL_YEARS"),
        "HOUSE_TAX_NAME": request.form.get("HOUSE_TAX_NAME"),
        "HOUSE_TAX_ISSUED_BY": request.form.get("HOUSE_TAX_ISSUED_BY"),

        "EC_DATE": request.form.get("EC_DATE"),
        "EC_NO": request.form.get("EC_NO"),

        # -------- ELECTRICITY --------
        "ELECTRICITY_BILL_DATE": request.form.get("ELECTRICITY_BILL_DATE"),
        "SERVICE_NO": request.form.get("SERVICE_NO"),
        "ELECTRICITY_NAME": request.form.get("ELECTRICITY_NAME"),

        # -------- MORTGAGE --------
        "MORTGAGE_DEED_NO": request.form.get("MORTGAGE_DEED_NO"),
        "MORTGAGE_DEED_DATE": request.form.get("MORTGAGE_DEED_DATE"),
        "MORTGAGE_COMPANY": request.form.get("MORTGAGE_COMPANY"),
    }

    # -------- BOOLEAN FLAGS --------
    context["HAS_ELECTRICITY_BILL"] = request.form.get("HAS_ELECTRICITY_BILL") == "true"
    context["HAS_MORTGAGE"] = request.form.get("HAS_MORTGAGE") == "true"

    # -------- DOCUMENTS (DYNAMIC) --------
    documents = []

    types = request.form.getlist("doc_type[]")
    numbers = request.form.getlist("doc_number[]")
    dates = request.form.getlist("doc_date[]")
    executants = request.form.getlist("doc_executant[]")
    owners = request.form.getlist("doc_owner[]")
    relations = request.form.getlist("doc_relation[]")
    worths = request.form.getlist("doc_worth[]")

    for i in range(len(types)):
        # STRICT VALIDATION (no empty junk rows)
        if types[i] and numbers[i] and dates[i]:
            documents.append({
                "type": types[i],
                "number": numbers[i],
                "date": dates[i],
                "executant": executants[i],
                "owner": owners[i],
                "relation": relations[i],
                "worth": worths[i]
            })

    context["DOCUMENTS"] = documents

    # -------- LOAD TEMPLATE --------
    doc = DocxTemplate("templates_docx/tyger_report.docx")
    doc.render(context)

    # -------- SAVE FILE --------
    filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    output_path = os.path.join("templates_docx", filename)
    doc.save(output_path)

    # -------- DOWNLOAD --------
    return send_file(output_path, as_attachment=True)


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(debug=True)
