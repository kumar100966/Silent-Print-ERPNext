import frappe, base64

# from frappe.utils.pdf import get_pdf,cleanup
from frappe import _


# Height calculation constants for thermal receipts (in mm)
# Must OVERESTIMATE — underestimates cause wkhtmltopdf pagination (visible gaps).
# crop_whitespace (PyMuPDF) trims bottom excess, so overestimation is safe.
RECEIPT_HEIGHT_CONFIG = {
    "header_height": 35,       # Company name (multi-line), WhatsApp, Call, Visit, TIN, separator,
                                # SALES RECEIPT bar, Receipt#, Date, Cashier, Sales Person
    "customer_height": 8,      # Customer name + bottom border
    "items_header_height": 8,  # Table header row (QTY ITEM PRICE AMT) + borders
    "item_height": 14,         # Base height per item (name line + UOM + dotted border)
    "chars_per_line": 20,      # ~20 chars fit in the ITEM column (not full 72mm width)
    "item_wrap_line": 4,       # Extra mm per wrapped line of item name
    "item_discount_height": 4, # Extra height for discount line
    "totals_height": 25,       # Subtotal + discount + VAT + grand total with double borders
    "payment_base_height": 10, # PAYMENT header + top border
    "payment_line_height": 6,  # Per payment method line
    "footer_height": 30,       # Thank you + pharmacy notice box + barcode + receipt-id + powered by
    "buffer_height": 15,       # Safety margin (covers rendering variance)
    "min_height": 100,         # Minimum receipt height
    "max_height": 3000,        # Thermal paper is continuous — no practical page limit
}


def calculate_receipt_height(doc):
    """
    Calculate the optimal receipt height based on document content.
    Supports POS Invoice (items/payments), POS Closing Entry
    (pos_transactions/payment_reconciliation/taxes), and other doctypes.
    Returns height in mm.

    Intentionally overestimates — crop_whitespace (PyMuPDF) trims excess.
    Underestimating causes wkhtmltopdf to paginate, creating visible gaps.
    """
    config = RECEIPT_HEIGHT_CONFIG
    chars_per_line = config["chars_per_line"]

    # Start with fixed sections
    height = config["header_height"] + config["customer_height"]

    # Detect document type and count line items accordingly
    items = (doc.get("items") or []) if doc else []
    pos_transactions = (doc.get("pos_transactions") or []) if doc else []
    taxes = (doc.get("taxes") or []) if doc else []
    payment_recon = (doc.get("payment_reconciliation") or []) if doc else []
    payments = (doc.get("payments") or []) if doc else []

    if pos_transactions:
        # POS Closing Entry: count transactions + tax lines + payment reconciliation rows
        height += len(pos_transactions) * config["item_height"]
        height += len(taxes) * config["payment_line_height"]
        # Payment reconciliation table (header + rows)
        if payment_recon:
            height += config["payment_base_height"] + (len(payment_recon) * config["payment_line_height"])
    else:
        # POS Invoice / Sales Invoice: count items + payments
        # Items table header row
        height += config["items_header_height"]

        # Per-item height with word-wrap estimation
        # The ITEM column is ~30mm wide (~20 chars at 10px Courier on 72mm receipt)
        item_height_total = 0
        for item in items:
            base = config["item_height"]
            name_len = len(item.get("item_name") or "")
            if name_len > chars_per_line:
                extra_lines = (name_len - 1) // chars_per_line
                base += extra_lines * config["item_wrap_line"]
            if item.get("discount_percentage") or item.get("discount_amount"):
                base += config["item_discount_height"]
            item_height_total += base
        height += item_height_total

        if payments:
            height += config["payment_base_height"] + (len(payments) * config["payment_line_height"])

    # Add totals section
    height += config["totals_height"]

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
    if not frappe.db.exists("Silent Print Format", silent_print_format):
        return

    silent_print_format_doc = frappe.get_doc("Silent Print Format", silent_print_format)

    html = frappe.get_print(doctype, name, silent_print_format, doc=doc, no_letterhead=no_letterhead)

    # Load the actual document for height calculation if auto_height is enabled
    actual_doc = None
    if silent_print_format_doc.get("auto_height"):
        try:
            actual_doc = frappe.get_doc(doctype, name) if not doc else doc
        except Exception:
            actual_doc = None

    options = get_pdf_options(silent_print_format_doc, actual_doc)
    pdf = get_pdf(html, options=options)

    print_type = silent_print_format_doc.default_print_type

    # For thermal receipts: convert PDF → ESC/POS raw bytes for direct printing.
    # This bypasses the Windows print spooler entirely, so the printer feeds
    # exactly as much paper as the content needs (true auto-height).
    #
    # pdf_to_escpos handles bottom whitespace trimming at the bitmap level
    # (no PDF structure modifications needed). This avoids MediaBox/CropBox
    # issues that caused header truncation with crop_pdf_whitespace.
    #
    # Size limit: WHB sends ESC/POS as a base64 JSON payload over WebSocket.
    # Java WebSocket implementations have message size limits (often 64-256 KB).
    # Large receipts exceed this. Fall back to PDF mode (uncropped) for those.
    MAX_RAW_BASE64_BYTES = 500_000  # ~500 KB — raised after fixing bitmap-level trim
    if print_type == "pos printer":
        try:
            raw_bytes = pdf_to_escpos(pdf)
            raw_base64 = base64.b64encode(raw_bytes).decode()
            if len(raw_base64) <= MAX_RAW_BASE64_BYTES:
                return {"raw_base64": raw_base64, "print_type": print_type}
            else:
                frappe.logger().info(
                    f"[SilentPrint] ESC/POS data too large ({len(raw_base64)} bytes > {MAX_RAW_BASE64_BYTES}), "
                    f"falling back to PDF for {doctype} {name}"
                )
        except Exception as e:
            frappe.log_error(
                title="ESC/POS Conversion Error",
                message=f"Falling back to PDF mode: {str(e)}\n{frappe.get_traceback()}",
            )
            # Fall through to PDF mode

    # PDF fallback: use the ORIGINAL uncropped PDF — WHB handles it correctly
    pdf_base64 = base64.b64encode(pdf)
    return {"pdf_base64": pdf_base64.decode(), "print_type": print_type}


def pdf_to_escpos(pdf_bytes, dpi=203):
    """
    Convert a PDF to ESC/POS raster commands for direct thermal printing.

    Renders the PDF page as a monochrome bitmap at the printer's native DPI,
    then wraps it in ESC/POS GS v 0 (raster bit image) commands.
    This bypasses the Windows print spooler — the printer feeds exactly
    as much paper as the content requires (true auto-height).

    Bottom whitespace is trimmed at the bitmap level — no PDF structure
    modifications (MediaBox/CropBox) needed. This avoids coordinate issues
    that can cause header truncation on some PDF generators.

    The raster is sent in bands of MAX_BAND_HEIGHT rows to stay within
    printer buffer limits.

    Args:
        pdf_bytes: Raw PDF file bytes
        dpi: Printer resolution (203 DPI is standard for 80mm thermal)

    Returns:
        bytes: Complete ESC/POS command sequence ready to send to printer
    """
    import fitz  # PyMuPDF

    MAX_BAND_HEIGHT = 255

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    # Render at printer's native DPI, grayscale
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
    width = pix.width
    height = pix.height
    stride = pix.stride
    samples = pix.samples

    doc.close()

    # --- Bitmap-level top whitespace trimming ---
    # Scan from top to find first row with any non-white pixel.
    # Handles wkhtmltopdf version differences that add top margin/padding.
    first_content_row = 0
    for y in range(height):
        row_start = y * stride
        row = samples[row_start : row_start + width]
        if min(row) < 250:  # Any pixel darker than near-white
            first_content_row = y
            break

    # --- Bitmap-level bottom whitespace trimming ---
    # Scan from bottom to find last row with any non-white pixel.
    # This replaces PDF-level cropping (crop_pdf_whitespace) which
    # modified MediaBox/CropBox and caused header truncation issues.
    last_content_row = height - 1
    for y in range(height - 1, -1, -1):
        row_start = y * stride
        row = samples[row_start : row_start + width]
        if min(row) < 250:  # Any pixel darker than near-white
            last_content_row = y
            break

    # Content = first_content_row..last_content_row + small buffer
    # Bottom buffer: ~2mm at 203 DPI = 16 pixels
    content_end = min(height, last_content_row + 16)
    content_height = content_end - first_content_row
    top_trimmed = first_content_row
    bottom_trimmed = height - content_end

    if top_trimmed > 0 or bottom_trimmed > 0:
        frappe.logger().info(
            f"[ESC/POS] Bitmap {width}x{height}px, content rows {first_content_row}..{last_content_row}, "
            f"trimmed {top_trimmed}px top + {bottom_trimmed}px bottom → {content_height}px"
        )

    # Convert grayscale to 1-bit monochrome (threshold at 128)
    # Each byte = 8 pixels, MSB first, 1=black 0=white
    # Only process rows first_content_row..content_end (skip top+bottom whitespace)
    byte_width = (width + 7) // 8
    raster = bytearray(byte_width * content_height)

    for y in range(content_height):
        row_offset = (y + first_content_row) * stride
        out_offset = y * byte_width
        for x in range(width):
            if samples[row_offset + x] < 128:  # Dark pixel
                raster[out_offset + (x >> 3)] |= (0x80 >> (x & 7))

    # Build ESC/POS command sequence
    escpos = bytearray()

    # 1. Initialize printer
    escpos.extend(b'\x1B\x40')  # ESC @ — reset printer

    # 2. Print raster in bands to avoid printer buffer overflow.
    #    Each band is a separate GS v 0 command.
    #    Format: 1D 76 30 m xL xH yL yH d1...dk
    #    m=0 (normal), xL/xH = bytes per row, yL/yH = number of rows
    rows_sent = 0
    while rows_sent < content_height:
        band_height = min(MAX_BAND_HEIGHT, content_height - rows_sent)

        escpos.extend(b'\x1D\x76\x30\x00')  # GS v 0, mode=0 (normal)
        escpos.append(byte_width & 0xFF)             # xL
        escpos.append((byte_width >> 8) & 0xFF)      # xH
        escpos.append(band_height & 0xFF)             # yL
        escpos.append((band_height >> 8) & 0xFF)      # yH

        # Slice the raster data for this band
        band_start = rows_sent * byte_width
        band_end = (rows_sent + band_height) * byte_width
        escpos.extend(raster[band_start:band_end])

        rows_sent += band_height

    # 3. Feed a few lines after content for readability
    escpos.extend(b'\x1B\x64\x04')  # ESC d 4 — feed 4 lines

    # 4. Partial cut (leave a small strip attached)
    escpos.extend(b'\x1D\x56\x42\x00')  # GS V 66 0 — partial cut

    return bytes(escpos)


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
    Crop bottom whitespace from thermal receipt PDFs.
    Uses pixel-based detection (renders as grayscale image) to find the last
    visible content row, then trims the page height down to content + small buffer.
    Preserves the top of the page as-is (avoids MediaBox origin changes that
    confuse some thermal printer pipelines).
    Falls back to original PDF on failure.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        frappe.log_error(title="Silent Print Crop Error", message="PyMuPDF (fitz) not installed. PDF cropping disabled.")
        return pdf_bytes

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        for page_num in range(len(doc)):
            page = doc[page_num]
            rect = page.rect

            # Render page as grayscale at 150 DPI for reliable text detection
            pix = page.get_pixmap(dpi=150, colorspace=fitz.csGRAY)
            w, h = pix.width, pix.height
            stride = pix.stride
            samples = pix.samples  # 1 byte per pixel, grayscale

            scale_y = rect.height / h if h > 0 else 1

            # Scan from bottom: find last row with any non-white pixel
            last_row = None
            for y in range(h - 1, -1, -1):
                row = samples[y * stride : y * stride + w]
                if min(row) < 250:
                    last_row = y
                    break

            if last_row is None:
                frappe.log_error(
                    title="Silent Print Crop Debug",
                    message=f"Page {page_num}: NO content detected (h={h}px)",
                )
                continue

            # Convert pixel row to PDF points
            content_bottom = (last_row + 1) * scale_y + rect.y0
            bottom_ws = rect.y1 - content_bottom

            # Only crop if bottom whitespace exceeds 3mm (~8.5pt)
            if bottom_ws <= 8.5:
                continue

            # Keep page origin (y0) unchanged — only reduce y1 (bottom edge)
            # 5pt (~1.8mm) buffer below last content pixel
            crop_y1 = min(rect.y1, content_bottom + 5)
            crop_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, crop_y1)

            page.set_mediabox(crop_rect)
            # CropBox inherits from MediaBox when not set explicitly.
            # Calling set_cropbox after set_mediabox can raise
            # "CropBox not in MediaBox" on some PyMuPDF versions, so
            # just delete any existing CropBox entry and let it inherit.
            try:
                page.set_cropbox(crop_rect)
            except ValueError:
                # CropBox validation failed — remove it so it defaults to MediaBox
                xref = page.xref
                doc.xref_set_key(xref, "CropBox", "null")

            cropped_mm = crop_rect.height / 2.835
            original_mm = rect.height / 2.835
            trimmed_mm = original_mm - cropped_mm
            frappe.log_error(
                title="Silent Print Crop Debug",
                message=f"Page {page_num}: {original_mm:.0f}mm -> {cropped_mm:.0f}mm (trimmed {trimmed_mm:.0f}mm bottom)",
            )

        output = doc.tobytes()
        doc.close()
        return output

    except Exception as e:
        frappe.log_error(
            title="Silent Print Crop Error",
            message=f"PDF cropping FAILED: {str(e)}\n{frappe.get_traceback()}",
        )
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
