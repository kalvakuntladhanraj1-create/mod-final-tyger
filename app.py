from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
from datetime import datetime
import os
import tempfile


app = Flask(__name__)


# ============================================================
# CONFIGURATION
# ============================================================

TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__),
    "templates_docx",
    "tyger_report.docx"
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def get_value(name, default=""):
    """
    Safely get a single form value.
    Empty/missing values become an empty string.
    """
    value = request.form.get(name, default)

    if value is None:
        return default

    return value.strip() if isinstance(value, str) else value


def format_date(value):
    """
    Convert HTML date:
        YYYY-MM-DD

    into:
        DD-MM-YYYY

    If empty or invalid, return empty string.
    """
    if not value:
        return ""

    value = value.strip()

    try:
        return datetime.strptime(
            value,
            "%Y-%m-%d"
        ).strftime("%d-%m-%Y")
    except (ValueError, TypeError):
        return value


def get_list(name):
    """
    Safely retrieve an HTML [] list field.
    """
    values = request.form.getlist(name)

    return [
        value.strip() if isinstance(value, str) else value
        for value in values
    ]


def value_at(values, index, default=""):
    """
    Safely retrieve a value from a list.

    If the list does not contain the requested index,
    return an empty value instead of crashing.
    """
    if index < len(values):
        value = values[index]

        if value is None:
            return default

        return value.strip() if isinstance(value, str) else value

    return default


def pop_next(values):
    """
    Take the next value from a list.

    Used for deed-specific fields because the HTML only
    creates those fields for the relevant deed type.

    Example:

        Gift -> relation field exists
        Sale -> relation field does not exist
        Gift -> relation field exists

    The relation list therefore contains only the two Gift
    values, not blank values for Sale.
    """
    if values:
        return values.pop(0)

    return ""


# ============================================================
# HOME
# ============================================================

@app.route("/")
def home():
    return render_template("sale.html")


# ============================================================
# GENERATE REPORT
# ============================================================

@app.route("/generate_sale", methods=["POST"])
def generate_sale():

    # ========================================================
    # BASIC INFORMATION
    # ========================================================

    context = {
        "DATE": format_date(get_value("DATE")),

        "APPLICATION_NO": get_value("APPLICATION_NO"),

        "APPLICANT_NAME": get_value("APPLICANT_NAME"),

        "APPLICANT_BOLD_OWNER": get_value(
            "APPLICANT_BOLD_OWNER"
        ),

        "LOAN_AMOUNT": get_value("LOAN_AMOUNT"),
    }


    # ========================================================
    # PROPERTY DETAILS
    # ========================================================

    context.update({

        "DOOR_NO": get_value("DOOR_NO"),

        "PLOT_NO": get_value("PLOT_NO"),

        "ASSESSMENT_NO": get_value(
            "ASSESSMENT_NO"
        ),

        # Kept because some older templates may use this
        "Assessment_No": get_value(
            "ASSESSMENT_NO"
        ),

        "EXTENT_YARDS": get_value(
            "EXTENT_YARDS"
        ),

        "SURVEY_NO": get_value(
            "SURVEY_NO"
        ),

        "ADDRESS": get_value(
            "ADDRESS"
        ),

        "VILLAGE": get_value(
            "VILLAGE"
        ),

        "GRAM_PANCHAYAT": get_value(
            "GRAM_PANCHAYAT"
        ),

        "MANDAL": get_value(
            "MANDAL"
        ),

        "SRO": get_value(
            "SRO"
        ),

        "RO": get_value(
            "RO"
        ),

        "DISTRICT": get_value(
            "DISTRICT"
        ),
    })


    # ========================================================
    # PROPERTY BOUNDARIES
    # ========================================================

    context.update({

        "EAST_BOUNDARY": get_value(
            "EAST_BOUNDARY"
        ),

        "WEST_BOUNDARY": get_value(
            "WEST_BOUNDARY"
        ),

        "NORTH_BOUNDARY": get_value(
            "NORTH_BOUNDARY"
        ),

        "SOUTH_BOUNDARY": get_value(
            "SOUTH_BOUNDARY"
        ),
    })


    # ========================================================
    # MEASUREMENTS
    # ========================================================

    context.update({

        "E_W": get_value("E_W"),

        "N_S": get_value("N_S"),

        "E_W_FEET": get_value(
            "E_W_FEET"
        ),

        "N_S_FEET": get_value(
            "N_S_FEET"
        ),

        "EXTENT_FEET": get_value(
            "EXTENT_FEET"
        ),
    })


    # ========================================================
    # PRIMARY DEED
    # ========================================================

    context.update({

        "DEED_NO": get_value(
            "DEED_NO"
        ),

        "DEED_DATE": format_date(
            get_value("DEED_DATE")
        ),
    })


    # ========================================================
    # PROPERTY TYPE
    #
    # DOCX:
    #
    # {% if Prop.type == "ancestral" %}
    #
    # {% elif Prop.type == "self" %}
    # ========================================================

    property_type = get_value(
        "PROPERTY_TYPE",
        "ancestral"
    ).lower()

    if property_type not in (
        "ancestral",
        "self"
    ):
        property_type = "ancestral"


    context["Prop"] = {
        "type": property_type
    }


    # ========================================================
    # ROOT OWNER / ROOT DOCUMENT
    #
    # These are needed by the self-acquired branch.
    #
    # They are still supplied even for ancestral property,
    # so a blank form never crashes.
    # ========================================================

    context["ROOT_OWNER"] = get_value(
        "ROOT_OWNER"
    )

    context["ROOT_DOC"] = {
        "number": get_value(
            "ROOT_DOC_NUMBER"
        ),

        "date": format_date(
            get_value("ROOT_DOC_DATE")
        )
    }


    # ========================================================
    # POSSESSION
    # ========================================================

    context.update({

        "POSSESSION_DATE": format_date(
            get_value("POSSESSION_DATE")
        ),

        "FROM_YEARS": get_value(
            "FROM_YEARS"
        ),

        "POSSESSION_NAME": get_value(
            "POSSESSION_NAME"
        ),
    })


    # ========================================================
    # HOUSE TAX
    # ========================================================

    context.update({

        "H_T_DATE": format_date(
            get_value("H_T_DATE")
        ),

        "HOUSE_TAX_RECIPT_NO": get_value(
            "HOUSE_TAX_RECIPT_NO"
        ),

        "FINANCIAL_YEARS": get_value(
            "FINANCIAL_YEARS"
        ),

        "HOUSE_TAX_NAME": get_value(
            "HOUSE_TAX_NAME"
        ),

        "HOUSE_TAX_ISSUED_BY": get_value(
            "HOUSE_TAX_ISSUED_BY"
        ),
    })


    # ========================================================
    # ENCUMBRANCE CERTIFICATE
    # ========================================================

    context.update({

        "EC_DATE": format_date(
            get_value("EC_DATE")
        ),

        "EC_NO": get_value(
            "EC_NO"
        ),
    })


    # ========================================================
    # ELECTRICITY BILL
    # ========================================================

    context.update({

        "HAS_ELECTRICITY_BILL":
            request.form.get(
                "HAS_ELECTRICITY_BILL"
            ) == "true",

        "ELECTRICITY_BILL_DATE":
            format_date(
                get_value(
                    "ELECTRICITY_BILL_DATE"
                )
            ),

        "SERVICE_NO":
            get_value(
                "SERVICE_NO"
            ),

        "ELECTRICITY_NAME":
            get_value(
                "ELECTRICITY_NAME"
            ),
    })


    # ========================================================
    # MORTGAGE
    # ========================================================

    context.update({

        "HAS_MORTGAGE":
            request.form.get(
                "HAS_MORTGAGE"
            ) == "true",

        "MORTGAGE_DEED_NO":
            get_value(
                "MORTGAGE_DEED_NO"
            ),

        "MORTGAGE_DEED_DATE":
            format_date(
                get_value(
                    "MORTGAGE_DEED_DATE"
                )
            ),

        "MORTGAGE_COMPANY":
            get_value(
                "MORTGAGE_COMPANY"
            ),
    })


    # ========================================================
    # DOCUMENTS
    # ========================================================

    types = get_list(
        "doc_type[]"
    )

    numbers = get_list(
        "doc_number[]"
    )

    dates = get_list(
        "doc_date[]"
    )

    executants = get_list(
        "doc_executant[]"
    )

    owners = get_list(
        "doc_owner[]"
    )

    relations = get_list(
        "doc_relation[]"
    )

    worths = get_list(
        "doc_worth[]"
    )

    type_rectified = get_list(
        "doc_type_rectified[]"
    )

    mistake_mentioned = get_list(
        "doc_mistake_mentioned[]"
    )

    schedule_items = get_list(
        "doc_schedule_item[]"
    )

    amount_paid = get_list(
        "doc_amount_paid[]"
    )

    favours = get_list(
        "doc_favour[]"
    )


    # ========================================================
    # IMPORTANT:
    #
    # Deed-specific HTML fields do NOT all have the same
    # number of entries.
    #
    # Example:
    #
    # Sale
    # Gift
    # Sale
    # Gift
    #
    # relation[] contains only:
    #
    # Gift relation
    # Gift relation
    #
    # Therefore we consume those lists only when the
    # corresponding deed type is encountered.
    # ========================================================

    documents = []


    for index, doc_type in enumerate(types):

        doc_type = (
            doc_type.strip()
            if isinstance(doc_type, str)
            else ""
        )


        # ----------------------------------------------------
        # COMMON FIELDS
        # ----------------------------------------------------

        document_number = value_at(
            numbers,
            index
        )

        document_date_raw = value_at(
            dates,
            index
        )

        document_date = format_date(
            document_date_raw
        )

        executant = value_at(
            executants,
            index
        )

        owner = value_at(
            owners,
            index
        )

        worth = value_at(
            worths,
            index
        )


        # ----------------------------------------------------
        # DOCUMENT OBJECT
        #
        # All possible fields are created so Jinja can safely
        # access them regardless of deed type.
        # ----------------------------------------------------

        document = {

            "type": doc_type,

            "number": document_number,

            "date": document_date,

            "executant": executant,

            "owner": owner,

            "relation": "",

            "worth": worth,

            "type_rectified": "",

            "mistake_mentioned": "",

            "schedule_item": "",

            "amount_paid": "",

            "favour": "",
        }


        # ----------------------------------------------------
        # GIFT
        # ----------------------------------------------------

        if doc_type == "Gift":

            document["relation"] = pop_next(
                relations
            )


        # ----------------------------------------------------
        # RECTIFICATION
        # ----------------------------------------------------

        elif doc_type == "Rectification":

            document["type_rectified"] = pop_next(
                type_rectified
            )

            document["mistake_mentioned"] = pop_next(
                mistake_mentioned
            )


        # ----------------------------------------------------
        # PARTITION
        # ----------------------------------------------------

        elif doc_type == "Partition":

            document["schedule_item"] = pop_next(
                schedule_items
            )


        # ----------------------------------------------------
        # SALE AGREEMENT
        # ----------------------------------------------------

        elif doc_type == "SaleAgreement":

            document["amount_paid"] = pop_next(
                amount_paid
            )


        # ----------------------------------------------------
        # DEPOSIT OF TITLE
        # ----------------------------------------------------

        elif doc_type == "DepositofTitle":

            document["favour"] = pop_next(
                favours
            )


        # ----------------------------------------------------
        # LOD
        # ----------------------------------------------------

        elif doc_type == "LOD":

            document["favour"] = pop_next(
                favours
            )


        # ----------------------------------------------------
        # RELINQUISHMENT
        #
        # Uses common executant / owner / worth fields.
        # ----------------------------------------------------

        elif doc_type == "Relinquishment":

            pass


        # ----------------------------------------------------
        # SALE
        #
        # Uses common fields.
        # ----------------------------------------------------

        elif doc_type == "Sale":

            pass


        # ----------------------------------------------------
        # ADD DOCUMENT
        # ----------------------------------------------------

        documents.append(
            document
        )


    # ========================================================
    # SERVER-SIDE SORTING
    #
    # Oldest document = No. 1
    # Newest document = last number
    #
    # This protects against relying only on JavaScript.
    # ========================================================

    def document_sort_key(document):

        raw_date = document.get(
            "date",
            ""
        )

        if not raw_date:
            # Undated documents go to the bottom.
            return (
                1,
                datetime.max
            )

        try:

            parsed_date = datetime.strptime(
                raw_date,
                "%d-%m-%Y"
            )

            return (
                0,
                parsed_date
            )

        except ValueError:

            return (
                1,
                datetime.max
            )


    documents.sort(
        key=document_sort_key
    )


    # ========================================================
    # FINAL DOCUMENT LIST
    # ========================================================

    context["DOCUMENTS"] = documents


    # ========================================================
    # TEMPLATE EXISTENCE CHECK
    # ========================================================

    if not os.path.exists(TEMPLATE_PATH):

        return (
            f"Template file not found: "
            f"{TEMPLATE_PATH}",
            500
        )


    # ========================================================
    # LOAD DOCX
    # ========================================================

    try:

        doc = DocxTemplate(
            TEMPLATE_PATH
        )

    except Exception as exc:

        return (
            f"Could not open DOCX template: {exc}",
            500
        )


    # ========================================================
    # RENDER DOCX
    # ========================================================

    try:

        doc.render(
            context
        )

    except Exception as exc:

        # Print the error to Render logs.
        app.logger.exception(
            "DOCX rendering failed"
        )

        return (
            f"Error while rendering DOCX: {exc}",
            500
        )


    # ========================================================
    # SAVE GENERATED DOCUMENT
    #
    # Use /tmp because Render allows temporary writable storage.
    # ========================================================

    filename = (
        "Legal_Scrutiny_Report_"
        + datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        + ".docx"
    )

    output_path = os.path.join(
        tempfile.gettempdir(),
        filename
    )


    try:

        doc.save(
            output_path
        )

    except Exception as exc:

        app.logger.exception(
            "Could not save generated DOCX"
        )

        return (
            f"Could not save generated report: {exc}",
            500
        )


    # ========================================================
    # DOWNLOAD
    # ========================================================

    return send_file(
        output_path,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


# ============================================================
# LOCAL DEVELOPMENT
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        ),
        debug=True
    )
