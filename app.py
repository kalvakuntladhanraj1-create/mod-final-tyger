from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
from datetime import datetime
import os
import tempfile

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__),
    "templates_docx",
    "tyger_report.docx"
)


def get_value(name, default=""):
    value = request.form.get(name, default)
    if value is None:
        return default
    return value.strip() if isinstance(value, str) else value


def format_date(value):
    if not value:
        return ""
    value = value.strip()
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%d-%m-%Y")
    except (ValueError, TypeError):
        return value


def get_list(name):
    values = request.form.getlist(name)
    return [
        value.strip() if isinstance(value, str) else value
        for value in values
    ]


def value_at(values, index, default=""):
    if index < len(values):
        value = values[index]
        if value is None:
            return default
        return value.strip() if isinstance(value, str) else value
    return default


def build_documents():
    types = get_list("doc_type[]")
    numbers = get_list("doc_number[]")
    dates = get_list("doc_date[]")
    executants = get_list("doc_executant[]")
    owners = get_list("doc_owner[]")
    relations = get_list("doc_relation[]")
    worths = get_list("doc_worth[]")
    type_rectified = get_list("doc_type_rectified[]")
    mistake_mentioned = get_list("doc_mistake_mentioned[]")
    schedule_items = get_list("doc_schedule_item[]")
    amount_paid = get_list("doc_amount_paid[]")
    favours = get_list("doc_favour[]")

    documents = []

    for index, doc_type in enumerate(types):
        doc_type = doc_type.strip() if isinstance(doc_type, str) else ""

        document = {
            "type": doc_type,
            "number": value_at(numbers, index),
            "date": format_date(value_at(dates, index)),
            "executant": value_at(executants, index),
            "owner": value_at(owners, index),
            "relation": value_at(relations, index),
            "worth": value_at(worths, index),
            "type_rectified": value_at(type_rectified, index),
            "mistake_mentioned": value_at(mistake_mentioned, index),
            "schedule_item": value_at(schedule_items, index),
            "amount_paid": value_at(amount_paid, index),
            "favour": value_at(favours, index),
        }

        documents.append(document)

    def document_sort_key(document):
        raw_date = document.get("date", "")

        if not raw_date:
            return (1, datetime.max)

        try:
            return (
                0,
                datetime.strptime(raw_date, "%d-%m-%Y")
            )
        except ValueError:
            return (1, datetime.max)

    documents.sort(key=document_sort_key)

    return documents


@app.route("/")
def home():
    return render_template("sale.html")


@app.route("/generate_sale", methods=["POST"])
def generate_sale():

    context = {
        "DATE": format_date(get_value("DATE")),
        "APPLICATION_NO": get_value("APPLICATION_NO"),
        "APPLICANT_NAME": get_value("APPLICANT_NAME"),
        "APPLICANT_BOLD_OWNER": get_value("APPLICANT_BOLD_OWNER"),
        "LOAN_AMOUNT": get_value("LOAN_AMOUNT"),

        "DOOR_NO": get_value("DOOR_NO"),
        "PLOT_NO": get_value("PLOT_NO"),
        "ASSESSMENT_NO": get_value("ASSESSMENT_NO"),
        "Assessment_No": get_value("ASSESSMENT_NO"),
        "EXTENT_YARDS": get_value("EXTENT_YARDS"),
        "SURVEY_NO": get_value("SURVEY_NO"),
        "ADDRESS": get_value("ADDRESS"),
        "VILLAGE": get_value("VILLAGE"),
        "GRAM_PANCHAYAT": get_value("GRAM_PANCHAYAT"),
        "MANDAL": get_value("MANDAL"),
        "SRO": get_value("SRO"),
        "RO": get_value("RO"),
        "DISTRICT": get_value("DISTRICT"),

        "EAST_BOUNDARY": get_value("EAST_BOUNDARY"),
        "WEST_BOUNDARY": get_value("WEST_BOUNDARY"),
        "NORTH_BOUNDARY": get_value("NORTH_BOUNDARY"),
        "SOUTH_BOUNDARY": get_value("SOUTH_BOUNDARY"),

        "E_W": get_value("E_W"),
        "N_S": get_value("N_S"),
        "E_W_FEET": get_value("E_W_FEET"),
        "N_S_FEET": get_value("N_S_FEET"),
        "EXTENT_FEET": get_value("EXTENT_FEET"),

        "DEED_NO": get_value("DEED_NO"),
        "DEED_DATE": format_date(get_value("DEED_DATE")),

        "ROOT_OWNER": get_value("ROOT_OWNER"),
        "ROOT_DOC": {
            "number": get_value("ROOT_DOC_NUMBER"),
            "date": format_date(get_value("ROOT_DOC_DATE"))
        },

        "POSSESSION_DATE": format_date(get_value("POSSESSION_DATE")),
        "FROM_YEARS": get_value("FROM_YEARS"),
        "POSSESSION_NAME": get_value("POSSESSION_NAME"),

        "H_T_DATE": format_date(get_value("H_T_DATE")),
        "HOUSE_TAX_RECIPT_NO": get_value("HOUSE_TAX_RECIPT_NO"),
        "FINANCIAL_YEARS": get_value("FINANCIAL_YEARS"),
        "HOUSE_TAX_NAME": get_value("HOUSE_TAX_NAME"),
        "HOUSE_TAX_ISSUED_BY": get_value("HOUSE_TAX_ISSUED_BY"),

        "EC_DATE": format_date(get_value("EC_DATE")),
        "EC_NO": get_value("EC_NO"),

        "HAS_ELECTRICITY_BILL":
            request.form.get("HAS_ELECTRICITY_BILL") == "true",
        "ELECTRICITY_BILL_DATE":
            format_date(get_value("ELECTRICITY_BILL_DATE")),
        "SERVICE_NO": get_value("SERVICE_NO"),
        "ELECTRICITY_NAME": get_value("ELECTRICITY_NAME"),

        "HAS_MORTGAGE":
            request.form.get("HAS_MORTGAGE") == "true",
        "MORTGAGE_DEED_NO": get_value("MORTGAGE_DEED_NO"),
        "MORTGAGE_DEED_DATE":
            format_date(get_value("MORTGAGE_DEED_DATE")),
        "MORTGAGE_COMPANY": get_value("MORTGAGE_COMPANY"),

        "DOCUMENTS": build_documents(),
    }

    property_type = get_value(
        "PROPERTY_TYPE",
        "ancestral"
    ).lower()

    if property_type not in ("ancestral", "self"):
        property_type = "ancestral"

    context["Prop"] = {"type": property_type}

    if not os.path.exists(TEMPLATE_PATH):
        return (
            f"Template file not found: {TEMPLATE_PATH}",
            500
        )

    try:
        doc = DocxTemplate(TEMPLATE_PATH)
        doc.render(context)
    except Exception as exc:
        app.logger.exception("DOCX rendering failed")
        return f"Error while rendering DOCX: {exc}", 500

    filename = (
        "Legal_Scrutiny_Report_"
        + datetime.now().strftime("%Y%m%d_%H%M%S")
        + ".docx"
    )

    output_path = os.path.join(
        tempfile.gettempdir(),
        filename
    )

    try:
        doc.save(output_path)
    except Exception as exc:
        app.logger.exception("Could not save generated DOCX")
        return f"Could not save generated report: {exc}", 500

    return send_file(
        output_path,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
        debug=True
    )
