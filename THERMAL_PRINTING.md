# Thermal Receipt Printing — ESC/POS Raw Mode

## Overview

Thermal receipt printing bypasses the Windows print spooler entirely by converting PDFs to ESC/POS raster commands server-side. The printer receives raw bitmap data and feeds exactly as much paper as the content requires (true auto-height).

**Printer**: Munbyn POS-80C (or compatible 80mm thermal)
**Paper**: CONTINUOUS_72 (72mm x 1000mm)
**Resolution**: 203 DPI
**Bridge**: Webapp Hardware Bridge (WHB)

## Architecture

```
Backend (create_pdf)
  │
  ├─ wkhtmltopdf: HTML → PDF (72mm wide, auto-calculated height)
  │
  ├─ pdf_to_escpos(): PDF → bitmap → ESC/POS raster (small/medium receipts)
  │   └─ Bitmap-level bottom whitespace trimming
  │   └─ GS v 0 raster in 255-row bands
  │   └─ Returns raw_base64
  │
  └─ PDF fallback: original uncropped PDF (if ESC/POS exceeds size limit)
      └─ Returns pdf_base64

Frontend (useSilentPrint.js)
  │
  ├─ raw_base64 → { type: "pos printer", raw_content: "<base64>" } → WHB
  └─ pdf_base64 → { type: "pos printer", file_content: "<base64>" } → WHB
```

## Key Files

| File | Purpose |
|------|---------|
| `silent_print/utils/print_format.py` | Core: PDF generation, ESC/POS conversion, height calculation |

## How It Works

### 1. PDF Generation (`create_pdf`)

wkhtmltopdf renders the print format HTML to PDF with:
- **Custom page size**: 72mm width, auto-calculated height
- **Zero margins**: Content fills the full receipt width
- **203 DPI**: Matches the thermal printer's native resolution

Height is calculated by `calculate_receipt_height()` which counts items, payments, and text wrapping. It intentionally **overestimates** — bottom whitespace is trimmed later at the bitmap level.

### 2. ESC/POS Conversion (`pdf_to_escpos`)

1. **Render**: PyMuPDF renders the PDF page as a grayscale bitmap at 203 DPI
2. **Trim**: Scans from bottom to find last row with content, discards blank rows
3. **Threshold**: Converts grayscale to 1-bit monochrome (pixel < 128 = black)
4. **Band**: Splits raster into 255-row bands (printer buffer compatibility)
5. **Command**: Wraps each band in a `GS v 0` raster bit image command

### 3. ESC/POS Command Sequence

```
ESC @           (0x1B 0x40)         — Initialize/reset printer
GS v 0          (0x1D 0x76 0x30 0x00) — Raster bit image, mode=0
  xL xH         — Bytes per row (low/high)
  yL yH         — Number of rows in band (low/high, max 255)
  d1..dk        — Raster data (1 bit/pixel, MSB first, black=1)
[repeat GS v 0 for each band]
ESC d 4         (0x1B 0x64 0x04)    — Feed 4 lines
GS V B 0        (0x1D 0x56 0x42 0x00) — Partial cut
```

### 4. Size Limit & PDF Fallback

ESC/POS data is sent as base64 in a JSON WebSocket message. If the base64 exceeds `MAX_RAW_BASE64_BYTES` (500 KB), the system falls back to sending the original uncropped PDF via WHB's standard PDF rendering path.

## Height Calculation (`RECEIPT_HEIGHT_CONFIG`)

Constants for estimating receipt content height (in mm):

| Section | Height | Notes |
|---------|--------|-------|
| Header | 35mm | Company name, contact info, receipt bar |
| Customer | 8mm | Customer name + border |
| Items header | 8mm | Table column headers |
| Per item | 14mm base | + 4mm per wrapped line, + 4mm if discount |
| Totals | 25mm | Subtotal, VAT, grand total |
| Payment header | 10mm | "PAYMENT" section |
| Per payment | 6mm | One line per method |
| Footer | 30mm | Thank you, barcode, powered-by |
| Buffer | 15mm | Safety margin |

**Min**: 100mm | **Max**: 3000mm

> **NEVER reduce these values** — underestimates cause wkhtmltopdf to paginate, creating visible gaps in the receipt.

## Important Lessons

### Bitmap-Level Trimming (not PDF-Level)

**NEVER modify PDF MediaBox/CropBox for thermal printing.** PyMuPDF's `set_mediabox()` changes the PDF coordinate system in ways that interact badly with wkhtmltopdf-generated PDFs, causing header truncation.

Instead, trim whitespace at the bitmap level:
1. Render the full PDF page as a bitmap
2. Scan from bottom to find last content row
3. Only convert content rows to ESC/POS

This is reliable because it operates on the actual rendered pixels, not on PDF structure.

### GS v 0 Banding

Never send a single `GS v 0` command with more than 255 rows. Many thermal printers have internal buffer limits. Split the raster into bands of 255 rows, each with its own `GS v 0` command.

### PDF Fallback Must Use Uncropped PDF

When falling back to WHB's PDF rendering mode, always send the **original uncropped PDF**. Cropped PDFs have modified dimensions that interact with WHB's `SHRINK_TO_FIT` scaling and the printer driver's paper configuration, causing content to be shrunk to unreadable sizes.

### ESC @ Reset

The `ESC @` command resets the printer to default settings. Always send it before raster commands to ensure consistent state.

### wkhtmltopdf Options for Thermal

Critical options that must not be overridden by HTML meta tags:
- `page-width: 72mm` — Must match printer paper width
- `margin-*: 0mm` — No margins for thermal receipts
- `dpi: 203` — Match printer native resolution
- `disable-smart-shrinking` — Prevent wkhtmltopdf from auto-scaling

The `_is_thermal` flag in `get_pdf_options()` protects these from being overridden by `read_options_from_html()`.

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|-------------|-----|
| Header truncated | PDF MediaBox/CropBox modification | Use bitmap-level trimming only |
| White space at bottom | Height overestimation | Expected — bitmap trim removes it |
| Content shrunk to tiny size | Cropped PDF sent via WHB PDF mode | Use original uncropped PDF for fallback |
| Top truncated on large bills | GS v 0 single command too large | Use 255-row banding |
| Nothing prints | ESC/POS data exceeds WHB WebSocket limit | Lower MAX_RAW_BASE64_BYTES |
| Pagination gaps | RECEIPT_HEIGHT_CONFIG too low | Increase height constants |
