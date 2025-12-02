import frappe, base64

# from frappe.utils.pdf import get_pdf,cleanup
from frappe import _


# Height calculation constants for thermal receipts (in mm)
RECEIPT_HEIGHT_CONFIG = {
    "header_height": 10,  # Company name, address, receipt info
    "customer_height": 8,  # Customer info section
    "item_height": 8,  # Height per item line (includes UOM, discount if any)
    "totals_height": 20,  # Subtotal, discount, tax, grand total
    "payment_base_height": 10,  # Payment section header
    "payment_line_height": 5,  # Per payment method
    "footer_height": 12,  # Thank you message
    "buffer_height": 5,  # Extra buffer for safety
    "min_height": 60,  # Minimum receipt height
    "max_height": 500,  # Maximum receipt height (prevents runaway)
}


def calculate_receipt_height(doc):
    """
    Calculate the optimal receipt height based on document content.
    Returns height in mm.
    """
    config = RECEIPT_HEIGHT_CONFIG

    # Start with fixed sections
    height = config["header_height"] + config["customer_height"]

    # Add item heights
    items_count = len(doc.get("items", [])) if doc else 0
    # Add extra height for items with discounts (they show an extra line)
    items_with_discount = sum(1 for item in (doc.get("items", []) if doc else []) if item.get("discount_percentage"))
    height += (items_count * config["item_height"]) + (items_with_discount * 2)

    # Add totals section
    height += config["totals_height"]

    # Add payment section if payments exist
    payments_count = len(doc.get("payments", [])) if doc else 0
    if payments_count > 0:
        height += config["payment_base_height"] + (payments_count * config["payment_line_height"])

    # Add footer
    height += config["footer_height"]

    # Add buffer
    height += config["buffer_height"]

    # Clamp to min/max
    height = max(config["min_height"], min(height, config["max_height"]))

    return int(height)


@frappe.whitelist()
def print_silently(doctype, name, print_format, print_type):
    user = frappe.db.get_single_value("Silent Print Settings", "print_user")
    tab_id = frappe.db.get_single_value("Silent Print Settings", "tab_id")
    pdf = create_pdf(doctype, name, print_format)
    data = {"doctype": doctype, "name": name, "print_format": print_format, "print_type": pdf["print_type"], "tab_id": tab_id, "pdf": pdf["pdf_base64"]}
    frappe.publish_realtime("print-silently", data, user=user)


@frappe.whitelist()
def set_master_tab(tab_id):
    query = 'update tabSingles set value={} where doctype="Silent Print Settings" and field="tab_id";'.format(tab_id)
    frappe.db.sql(query)
    frappe.publish_realtime("update_master_tab", {"tab_id": tab_id})


@frappe.whitelist()
def create_pdf(doctype, name, silent_print_format, doc=None, no_letterhead=0):
    html = frappe.get_print(doctype, name, silent_print_format, doc=doc, no_letterhead=no_letterhead)
    if not frappe.db.exists("Silent Print Format", silent_print_format):
        return

    silent_print_format_doc = frappe.get_doc("Silent Print Format", silent_print_format)

    # Load the actual document for height calculation if auto_height is enabled
    actual_doc = None
    if silent_print_format_doc.get("auto_height"):
        try:
            actual_doc = frappe.get_doc(doctype, name) if not doc else doc
        except Exception:
            actual_doc = None

    options = get_pdf_options(silent_print_format_doc, actual_doc)
    pdf = get_pdf(html, options=options)

    # Optionally crop the PDF to remove whitespace
    if silent_print_format_doc.get("crop_whitespace"):
        pdf = crop_pdf_whitespace(pdf)

    pdf_base64 = base64.b64encode(pdf)
    return {"pdf_base64": pdf_base64.decode(), "print_type": silent_print_format_doc.default_print_type}


def get_pdf_options(silent_print_format, doc=None):
    """
    Generate PDF options for wkhtmltopdf.
    If auto_height is enabled, calculates height based on document content.
    """
    options = {"page-size": silent_print_format.get("page_size") or "A4"}

    if silent_print_format.get("page_size") == "Custom":
        custom_width = silent_print_format.get("custom_width") or "80mm"

        # Determine height: auto-calculate or use configured value
        if silent_print_format.get("auto_height") and doc:
            calculated_height = calculate_receipt_height(doc)
            custom_height = f"{calculated_height}mm"
        else:
            custom_height = silent_print_format.get("custom_height") or "297mm"

        # Ensure units are specified
        if custom_width and not any(u in str(custom_width) for u in ["mm", "cm", "in", "px"]):
            custom_width = f"{custom_width}mm"
        if custom_height and not any(u in str(custom_height) for u in ["mm", "cm", "in", "px"]):
            custom_height = f"{custom_height}mm"

        options = {
            "page-width": custom_width,
            "page-height": custom_height,
            # Zero margins for thermal/receipt printing
            "margin-top": "0mm",
            "margin-bottom": "0mm",
            "margin-left": "0mm",
            "margin-right": "0mm",
            # Thermal receipt specific options
            "dpi": "203",  # Standard thermal printer DPI
            "zoom": "1",
            "no-pdf-compression": "",
            # Mark as thermal to prevent option overrides
            "_is_thermal": True,
        }
    return options


def crop_pdf_whitespace(pdf_bytes):
    """
    Crop whitespace from the bottom of a PDF.
    Requires PyMuPDF (fitz) to be installed.
    Falls back to original PDF if cropping fails.
    """
    try:
        import fitz  # PyMuPDF

        # Open PDF from bytes
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        for page_num in range(len(pdf_doc)):
            page = pdf_doc[page_num]

            # Get the bounding box of all content on the page
            # This finds the actual content boundaries
            blocks = page.get_text("dict")["blocks"]

            if not blocks:
                continue

            # Find the lowest point of actual content
            max_y = 0
            for block in blocks:
                if "bbox" in block:
                    max_y = max(max_y, block["bbox"][3])
                elif "lines" in block:
                    for line in block["lines"]:
                        if "bbox" in line:
                            max_y = max(max_y, line["bbox"][3])

            if max_y > 0:
                # Add a small buffer (5 points ~ 1.7mm)
                max_y += 15

                # Get current page dimensions
                rect = page.rect

                # Only crop if there's significant whitespace (more than 20 points)
                if rect.height - max_y > 20:
                    # Create new crop box
                    new_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, max_y)
                    page.set_cropbox(new_rect)

        # Save to bytes
        output = pdf_doc.tobytes()
        pdf_doc.close()
        return output

    except ImportError:
        # PyMuPDF not installed, return original
        frappe.log_error("PyMuPDF (fitz) not installed. PDF cropping disabled.", "Silent Print")
        return pdf_bytes
    except Exception as e:
        # Any other error, return original
        frappe.log_error(f"PDF cropping failed: {str(e)}", "Silent Print")
        return pdf_bytes


from distutils.version import LooseVersion
import pdfkit
import six
import io
from bs4 import BeautifulSoup

# PyPDF2 3.x compatibility
try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    # Fallback for older PyPDF2 versions
    from PyPDF2 import PdfFileReader as PdfReader, PdfFileWriter as PdfWriter
from frappe.utils import scrub_urls
from frappe.utils.pdf import get_file_data_from_writer, read_options_from_html, get_wkhtmltopdf_version

PDF_CONTENT_ERRORS = ["ContentNotFoundError", "ContentOperationNotPermittedError", "UnknownContentError", "RemoteHostClosedError"]


def get_pdf(html, options=None, output=None):
    html = scrub_urls(html)
    html, options = prepare_options(html, options)

    options.update({"disable-javascript": "", "disable-local-file-access": ""})

    filedata = ""
    if LooseVersion(get_wkhtmltopdf_version()) > LooseVersion("0.12.3"):
        options.update({"disable-smart-shrinking": ""})

    try:
        # Set filename property to false, so no file is actually created
        filedata = pdfkit.from_string(html, False, options=options or {})

        # Create in-memory binary streams from filedata and create a PdfReader object
        reader = PdfReader(io.BytesIO(filedata))
    except OSError as e:
        if any([error in str(e) for error in PDF_CONTENT_ERRORS]):
            if not filedata:
                frappe.throw(_("PDF generation failed because of broken image links"))

            # allow pdfs with missing images if file got created
            if output:  # output is a PdfWriter object
                for page in reader.pages:
                    output.add_page(page)
        else:
            raise

    if "password" in options:
        password = options["password"]
        if six.PY2:
            password = frappe.safe_encode(password)

    if output:
        for page in reader.pages:
            output.add_page(page)
        return output

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    if "password" in options:
        writer.encrypt(password)

    filedata = get_file_data_from_writer(writer)

    return filedata


def prepare_options(html, options):
    if not options:
        options = {}

    # Check if this is a thermal receipt (custom page size with zero margins)
    is_thermal = options.pop("_is_thermal", False)

    # Store thermal-specific settings before they can be overwritten
    thermal_settings = {}
    if is_thermal:
        thermal_settings = {
            "margin-top": options.get("margin-top", "0mm"),
            "margin-bottom": options.get("margin-bottom", "0mm"),
            "margin-left": options.get("margin-left", "0mm"),
            "margin-right": options.get("margin-right", "0mm"),
            "page-width": options.get("page-width"),
            "page-height": options.get("page-height"),
            "dpi": options.get("dpi"),
            "zoom": options.get("zoom"),
        }

    options.update(
        {
            "print-media-type": None,
            "background": None,
            "images": None,
            "quiet": None,
            "encoding": "UTF-8",
        }
    )

    # Only set default margins if not thermal and not already specified
    if not is_thermal:
        if not options.get("margin-right"):
            options["margin-right"] = "15mm"
        if not options.get("margin-left"):
            options["margin-left"] = "15mm"
        if not options.get("margin-top"):
            options["margin-top"] = "15mm"
        if not options.get("margin-bottom"):
            options["margin-bottom"] = "15mm"

    # Read options from HTML (but don't let them override thermal settings)
    html, html_options = read_options_from_html(html)
    if html_options:
        if is_thermal:
            # For thermal receipts, only apply non-margin HTML options
            for key, value in html_options.items():
                if "margin" not in key.lower() and "page" not in key.lower():
                    options[key] = value
        else:
            options.update(html_options)

    # Restore thermal settings (ensures they're not overridden)
    if is_thermal:
        options.update(thermal_settings)

    # cookies
    if frappe.session and frappe.session.sid:
        options["cookie"] = [("sid", "{0}".format(frappe.session.sid))]

    return html, options
