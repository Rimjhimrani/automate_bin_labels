import streamlit as st
import pandas as pd
import os
from reportlab.lib.pagesizes import landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak, Image
from reportlab.lib.units import cm, inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.utils import ImageReader
from io import BytesIO
import subprocess
import sys
import re
import tempfile

# Define sticker dimensions
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# Define content box dimensions
CONTENT_BOX_WIDTH = 10 * cm
CONTENT_BOX_HEIGHT = 7.2 * cm

# ── PIL ──────────────────────────────────────────────────────────────────────
try:
    from PIL import Image as PILImage
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pillow'])
    from PIL import Image as PILImage

# ── qrcode ───────────────────────────────────────────────────────────────────
try:
    import qrcode
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'qrcode'])
    import qrcode

# ── Shared paragraph styles ───────────────────────────────────────────────────
bold_style = ParagraphStyle(name='Bold',        fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, leading=14)
desc_style = ParagraphStyle(name='Description', fontName='Helvetica',      fontSize=11, alignment=TA_CENTER, leading=12)
qty_style  = ParagraphStyle(name='Quantity',    fontName='Helvetica',      fontSize=11, alignment=TA_CENTER, leading=12)


# ════════════════════════════════════════════════════════════════════════════════
# AUTO-DETECT MODEL COLUMNS  (replaces all the old hardcoded logic)
# ════════════════════════════════════════════════════════════════════════════════

# Column-name fragments that identify NON-model columns
_NON_MODEL = [
    'PART', 'DESC', 'NAME', 'QTY', 'QUANTITY', 'BIN', 'LOC',
    'STATION', 'RACK', 'LEVEL', 'CELL', 'ABB', 'STORE', 'FLOOR',
    'ZONE', 'POSITION', 'NO', 'NUM', 'VEH', 'CAR', 'MODEL',
    'BUS', 'VEHICLE', 'TYPE',
]

def detect_model_columns(original_columns):
    """
    Scan the Excel header and return exactly up to 4 column names
    whose headers look like bus-model labels (not Part No, Desc, Qty, location, etc.).
    The column header text itself becomes the label shown in the MTM box.
    """
    model_cols = []
    for col in original_columns:
        col_str = str(col).strip()
        if not col_str or col_str.lower().startswith('unnamed:') or col_str.lower() == 'nan':
            continue
        col_upper = col_str.upper()
        if any(p in col_upper for p in _NON_MODEL):
            continue
        model_cols.append(col)
        if len(model_cols) == 4:
            break
    return model_cols


def get_row_model_quantities(row, model_cols):
    """
    Return an ordered dict {model_label: qty_string} for every model column.
    Cleans up float formatting (10.0 → 10), zero-suppresses, NaN-suppresses.
    """
    result = {}
    for col in model_cols:
        label = str(col).strip()
        val   = row.get(col, "")
        if pd.isna(val) or str(val).strip() in ('', 'nan'):
            result[label] = ''
            continue
        try:
            f = float(val)
            result[label] = '' if f == 0 else (str(int(f)) if f == int(f) else str(f))
        except (ValueError, TypeError):
            result[label] = str(val).strip()
    return result


# ════════════════════════════════════════════════════════════════════════════════
# QR CODE
# ════════════════════════════════════════════════════════════════════════════════

def generate_qr_code(data_string):
    try:
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4)
        qr.add_data(data_string)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        buf = BytesIO()
        qr_img.save(buf, format='PNG')
        buf.seek(0)
        return Image(buf, width=2.2*cm, height=2.2*cm)
    except Exception as e:
        st.error(f"Error generating QR code: {e}")
        return None


# ════════════════════════════════════════════════════════════════════════════════
# LOCATION HELPERS
# ════════════════════════════════════════════════════════════════════════════════

def parse_location_string(location_str):
    location_parts = [''] * 7
    if not location_str or not isinstance(location_str, str):
        return location_parts
    matches = re.findall(r'([^_\s]+)', location_str.strip())
    for i, m in enumerate(matches[:7]):
        location_parts[i] = m
    return location_parts


def extract_location_data_from_excel(row_data, active_model_label=''):
    """
    Returns 7 values for Line Location cells:
    [Active Bus Model, Station No, Rack, Rack 1st digit, Rack 2nd digit, Level, Cell]
    Slot 0 = the bus model that has a qty/veh value (passed in as active_model_label).
    Uses case-insensitive key matching.
    """
    if hasattr(row_data, 'to_dict'):
        raw = row_data.to_dict()
    else:
        raw = dict(row_data)
    upper_lookup = {str(k).upper(): v for k, v in raw.items()}

    def find_val(possible_names, default=''):
        for name in possible_names:
            v = upper_lookup.get(name.upper())
            if v is not None and pd.notna(v) and str(v).strip().lower() not in ('', 'nan'):
                raw_str = str(v).strip()
                try:
                    f = float(raw_str)
                    return str(int(f)) if f == int(f) else raw_str
                except (ValueError, TypeError):
                    return raw_str
        return default

    return [
        active_model_label,
        find_val(['Station No', 'Station_No', 'STATIONNO']),
        find_val(['Rack', 'RACK']),
        find_val(['Rack No (1st digit)', 'Rack_No_1st', 'RACK_NO_1ST', 'Rack No 1st digit']),
        find_val(['Rack No (2nd digit)', 'Rack_No_2nd', 'RACK_NO_2ND', 'Rack No 2nd digit']),
        find_val(['Level', 'LEVEL']),
        find_val(['Cell', 'CELL']),
    ]


def extract_store_location_data_from_excel(row_data):
    """Case-insensitive lookup for ABB store location columns."""
    if hasattr(row_data, 'to_dict'):
        raw = row_data.to_dict()
    else:
        raw = dict(row_data)
    upper_lookup = {str(k).upper(): v for k, v in raw.items()}

    def get(key, default=''):
        v = upper_lookup.get(key.upper())
        if v is None or not pd.notna(v) or str(v).strip().lower() in ('', 'nan'):
            return default
        raw_str = str(v).strip()
        try:
            f = float(raw_str)
            return str(int(f)) if f == int(f) else raw_str
        except (ValueError, TypeError):
            return raw_str

    return [
        get('ABB ZONE', ''),
        get('ABB LOCATION', ''),
        get('ABB FLOOR', ''),
        get('ABB RACK NO', ''),
        get('ABB LEVEL IN RACK', ''),
        get('ABB CELL', ''),
        get('ABB NO', ''),
    ]


# ════════════════════════════════════════════════════════════════════════════════
# MAIN PDF GENERATOR
# ════════════════════════════════════════════════════════════════════════════════

def generate_sticker_labels(excel_file_path, output_pdf_path, status_callback=None):

    def log(msg):
        if status_callback:
            status_callback(msg)
        else:
            st.write(msg)

    log(f"Processing file: {excel_file_path}")

    # ── border drawn on every page ───────────────────────────────────────────
    def draw_border(canvas, doc):
        canvas.saveState()
        x_off = (STICKER_WIDTH - CONTENT_BOX_WIDTH) / 2
        y_off = STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm
        canvas.setStrokeColor(colors.Color(0, 0, 0, alpha=0.95))
        canvas.setLineWidth(1.8)
        canvas.rect(x_off + doc.leftMargin, y_off, CONTENT_BOX_WIDTH - 0.2*cm, CONTENT_BOX_HEIGHT)
        canvas.restoreState()

    # ── read file ────────────────────────────────────────────────────────────
    try:
        if excel_file_path.lower().endswith('.csv'):
            df = pd.read_csv(excel_file_path, keep_default_na=False)
        else:
            try:
                df = pd.read_excel(excel_file_path, keep_default_na=False)
            except Exception:
                df = pd.read_excel(excel_file_path, engine='openpyxl', keep_default_na=False)
        log(f"✅ Read {len(df)} rows. Columns: {df.columns.tolist()}")
    except Exception as e:
        log(f"❌ Error reading file: {e}")
        return None

    # ── identify standard columns ────────────────────────────────────────────
    original_columns = df.columns.tolist()

    # Work with original casing for model detection, uppercase for standard cols
    df_upper = df.copy()
    df_upper.columns = [str(c).upper() for c in df_upper.columns]
    cols_upper = df_upper.columns.tolist()

    part_no_col = next((c for c in cols_upper if 'PART' in c and ('NO' in c or 'NUM' in c or '#' in c)),
                       next((c for c in cols_upper if c in ['PARTNO', 'PART']), cols_upper[0]))

    desc_col = next((c for c in cols_upper if 'DESC' in c),
                    next((c for c in cols_upper if 'NAME' in c), cols_upper[1] if len(cols_upper) > 1 else part_no_col))

    qty_bin_col = next((c for c in cols_upper if 'QTY/BIN' in c or 'QTY_BIN' in c or 'QTYBIN' in c),
                       next((c for c in cols_upper if 'QTY' in c and 'BIN' in c),
                            next((c for c in cols_upper if 'QTY' in c),
                                 next((c for c in cols_upper if 'QUANTITY' in c), None))))

    loc_col = next((c for c in cols_upper if 'LOC' in c or 'POS' in c or 'LOCATION' in c),
                   cols_upper[2] if len(cols_upper) > 2 else desc_col)

    store_loc_col = next((c for c in cols_upper if 'STORE' in c and 'LOC' in c),
                         next((c for c in cols_upper if 'STORELOCATION' in c), None))

    log(f"Part No: {part_no_col} | Desc: {desc_col} | Qty/Bin: {qty_bin_col} | Loc: {loc_col}")

    # ── AUTO-DETECT MODEL COLUMNS (from original casing) ────────────────────
    model_cols = detect_model_columns(original_columns)

    # Pad to exactly 4 slots; empty string = blank label
    while len(model_cols) < 4:
        model_cols.append(None)

    model_labels = [str(c).strip() if c else '' for c in model_cols]
    log(f"📦 Auto-detected model columns (up to 4): {[l for l in model_labels if l]}")

    # ── document setup ───────────────────────────────────────────────────────
    doc = SimpleDocTemplate(
        output_pdf_path, pagesize=STICKER_PAGESIZE,
        topMargin=0.2*cm,
        bottomMargin=(STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm),
        leftMargin=0.1*cm, rightMargin=0.1*cm
    )
    content_width = CONTENT_BOX_WIDTH - 0.2*cm
    all_elements  = []
    total_rows    = len(df)

    # ── per-row sticker ──────────────────────────────────────────────────────
    for index, (_, row_upper) in enumerate(df_upper.iterrows()):
        # Also get original-cased row for model qty lookup
        row_orig = df.iloc[index]

        if status_callback:
            status_callback(f"Creating sticker {index+1} of {total_rows} ({int((index+1)/total_rows*100)}%)")

        elements = []

        # ── basic fields ─────────────────────────────────────────────────────
        part_no  = str(row_upper.get(part_no_col, ""))
        desc     = str(row_upper.get(desc_col, ""))
        qty_bin  = str(row_upper.get(qty_bin_col, "")) if qty_bin_col else ""
        if qty_bin.lower() == 'nan':
            qty_bin = ""

        location_str    = str(row_upper.get(loc_col, "")) if loc_col else ""
        store_location  = str(row_upper.get(store_loc_col, "")) if store_loc_col else ""

        # ── model quantities from auto-detected columns ───────────────────────
        # Build dict {label: qty} using original-cased row
        row_orig_dict = row_orig.to_dict()
        mtm_quantities = {}
        for col in model_cols:
            if col is None:
                continue
            label = str(col).strip()
            val   = row_orig_dict.get(col, "")
            if pd.isna(val) or str(val).strip() in ('', 'nan'):
                mtm_quantities[label] = ''
                continue
            try:
                f = float(val)
                mtm_quantities[label] = '' if f == 0 else (str(int(f)) if f == int(f) else str(f))
            except (ValueError, TypeError):
                mtm_quantities[label] = str(val).strip()

        # ── QR code ──────────────────────────────────────────────────────────
        qty_veh_str = ", ".join(f"{k}:{v}" for k, v in mtm_quantities.items() if v)
        # Active model = first model that has a qty/veh (used in Line Location slot 0)
        active_model_label = next((k for k, v in mtm_quantities.items() if v), "")

        qr_data  = (f"Part No: {part_no}\nDescription: {desc}\n"
                    f"Location: {location_str}\nStore Location: {store_location}\n"
                    f"QTY/VEH: {qty_veh_str}\nQTY/BIN: {qty_bin}")
        qr_image = generate_qr_code(qr_data)

        # ── row heights ───────────────────────────────────────────────────────
        header_row_height   = 0.9*cm
        desc_row_height     = 1.0*cm
        qty_row_height      = 0.5*cm
        location_row_height = 0.5*cm

        # ── main table ────────────────────────────────────────────────────────
        main_table = Table([
            ["Part No",     Paragraph(part_no, bold_style)],
            ["Description", Paragraph(desc[:47] + "..." if len(desc) > 50 else desc, desc_style)],
            ["Qty/Bin",     Paragraph(qty_bin, qty_style)],
        ], colWidths=[content_width/3, content_width*2/3],
           rowHeights=[header_row_height, desc_row_height, qty_row_height])

        main_table.setStyle(TableStyle([
            ('GRID',     (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN',    (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN',   (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (0, -1),  'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (0, -1),  11),
        ]))
        elements.append(main_table)

        # ── store / line location ─────────────────────────────────────────────
        inner_table_width = content_width * 2 / 3
        col_proportions   = [1.5, 2.5, 0.7, 0.8, 0.8, 0.7, 0.9]
        total_prop        = sum(col_proportions)
        inner_col_widths  = [w * inner_table_width / total_prop for w in col_proportions]

        for label_text, values_fn in [
            ("Store Location", lambda r=row_orig_dict: extract_store_location_data_from_excel(r)),
            # Pass original-cased row so extract_location_data_from_excel builds its own uppercase lookup
            ("Line Location",  lambda r=row_orig.to_dict(), m=active_model_label: extract_location_data_from_excel(r, m)),
        ]:
            loc_values = values_fn()

            lbl = Paragraph(label_text, ParagraphStyle(
                name=label_text.replace(" ", ""), fontName='Helvetica-Bold',
                fontSize=11, alignment=TA_CENTER
            ))
            inner = Table([loc_values], colWidths=inner_col_widths, rowHeights=[location_row_height])
            inner.setStyle(TableStyle([
                ('GRID',     (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
                ('ALIGN',    (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN',   (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
            ]))
            outer = Table([[lbl, inner]],
                          colWidths=[content_width/3, inner_table_width],
                          rowHeights=[location_row_height])
            outer.setStyle(TableStyle([
                ('GRID',   (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
                ('ALIGN',  (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            elements.append(outer)

        elements.append(Spacer(1, 0.5*cm))

        # ── MTM boxes (always exactly 4) ──────────────────────────────────────
        mtm_box_width  = 1.55*cm
        mtm_row_height = 1.5*cm

        def mtm_label_style(name):
            return ParagraphStyle(name=name, fontName='Helvetica-Bold',
                                  fontSize=9, alignment=TA_CENTER, leading=10)

        def mtm_value_style(name):
            return ParagraphStyle(name=name, fontName='Helvetica-Bold',
                                  fontSize=10, alignment=TA_CENTER)

        header_row = []
        value_row  = []
        for i in range(4):
            lbl = model_labels[i]
            col = model_cols[i]
            # header cell
            header_row.append(Paragraph(lbl, mtm_label_style(f'Hdr{i}')))
            # value cell
            qty = mtm_quantities.get(lbl, '') if lbl else ''
            if qty:
                value_row.append(Paragraph(f"<b>{qty}</b>", mtm_value_style(f'Val{i}')))
            else:
                value_row.append("")

        mtm_table = Table(
            [header_row, value_row],
            colWidths=[mtm_box_width] * 4,
            rowHeights=[mtm_row_height/2, mtm_row_height/2]
        )
        mtm_table.setStyle(TableStyle([
            ('GRID',     (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN',    (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN',   (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0),  'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))

        # ── QR table ──────────────────────────────────────────────────────────
        qr_width  = 2.2*cm
        qr_height = 2.2*cm
        qr_cell   = qr_image if qr_image else Paragraph("QR", ParagraphStyle(
            name='QRPlaceholder', fontName='Helvetica-Bold', fontSize=12, alignment=TA_CENTER))
        qr_table  = Table([[qr_cell]], colWidths=[qr_width], rowHeights=[qr_height])
        qr_table.setStyle(TableStyle([
            ('ALIGN',  (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        # ── bottom layout: [gap | mtm boxes | gap | QR] ───────────────────────
        left_gap   = 0.1*cm
        mid_gap    = 0.2*cm
        right_gap  = content_width - (mtm_box_width * 4) - qr_width - left_gap - mid_gap

        bottom = Table(
            [["", mtm_table, "", qr_table]],
            colWidths=[left_gap, mtm_box_width * 4, right_gap, qr_width],
            rowHeights=[mtm_row_height]
        )
        bottom.setStyle(TableStyle([
            ('ALIGN',  (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(bottom)

        all_elements.extend(elements)
        if index < total_rows - 1:
            all_elements.append(PageBreak())

    # ── build PDF ─────────────────────────────────────────────────────────────
    try:
        doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
        log(f"✅ PDF created: {output_pdf_path}")
        return output_pdf_path
    except Exception as e:
        log(f"❌ Error creating PDF: {e}")
        import traceback; traceback.print_exc()
        return None


# ════════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ════════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(page_title="Chakan Bin Label Generator", layout="wide")
    st.title("🏷️ Chakan Bin Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px;'>"
        "Designed and Developed by Agilomatrix</p>", unsafe_allow_html=True
    )

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV file", type=['xlsx', 'xls', 'csv'],
        help="Select a file containing part numbers, descriptions, and location data"
    )

    if uploaded_file is not None:
        suffix = os.path.splitext(uploaded_file.name)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        st.success(f"File uploaded: {uploaded_file.name}")

        try:
            df_preview = (pd.read_csv(tmp_path) if uploaded_file.name.lower().endswith('.csv')
                          else pd.read_excel(tmp_path))
            st.subheader("📊 Data Preview")
            st.dataframe(df_preview.head(10))
            st.info(f"Total rows: {len(df_preview)}")

            # Show which model columns were auto-detected
            model_cols_detected = detect_model_columns(df_preview.columns.tolist())
            if model_cols_detected:
                st.success(f"🔍 Auto-detected model columns: **{', '.join(str(c) for c in model_cols_detected)}**")
            else:
                st.warning("⚠️ No model columns detected. Check your column headers.")
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return

        if st.button("🏷️ Generate Labels", type="primary"):
            output_filename = f"sticker_labels_{os.path.splitext(uploaded_file.name)[0]}.pdf"
            progress_bar = st.progress(0)
            status_text  = st.empty()

            def update_status(message):
                status_text.text(message)
                if "Creating sticker" in message and "of" in message:
                    try:
                        parts   = message.split()
                        current = int(parts[2])
                        total   = int(parts[4])
                        progress_bar.progress(current / total)
                    except Exception:
                        pass

            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                    pdf_path = generate_sticker_labels(tmp_path, tmp_pdf.name, update_status)

                if pdf_path:
                    with open(pdf_path, 'rb') as f:
                        pdf_data = f.read()
                    st.download_button(
                        label="📥 Download Sticker Labels PDF",
                        data=pdf_data, file_name=output_filename,
                        mime="application/pdf", type="primary"
                    )
                    st.success("✅ Sticker labels generated successfully!")
                    try:
                        os.unlink(tmp_path)
                        os.unlink(pdf_path)
                    except Exception:
                        pass
            except Exception as e:
                st.error(f"Error generating stickers: {e}")
                import traceback; st.code(traceback.format_exc())
    else:
        st.info("👆 Please upload an Excel or CSV file to get started")

    # ── sample format reference ───────────────────────────────────────────────
    st.subheader("📋 Reference For Data Format")
    sample_data = {
        'Part No':            ['08-DRA-14-02', 'P0012124-07', 'P0012126-07'],
        'Part Desc':          ['BELLOW ASSY. WITH RETAINING CLIP', 'GUARD RING (hirkesh)', 'GUARD RING SEAL (hirkesh)'],
        'Bin Type':           ['TOTE', 'BIN C', 'BIN A'],
        'Qty/bin':            [360, 20, 120],
        # ↓ These 4 columns become the MTM boxes automatically
        '135 KW':             [10, '', ''],
        '60 KW':              ['', 5, ''],
        'C':                  ['', '', 2],
        '4W':                 ['', '', ''],
        'Station No':         ['CW40RH', 'CW40RH', 'CW40RH'],
        'Rack':               ['R', 'R', 'R'],
        'Rack No (1st digit)':[ 0,  0,  0],
        'Rack No (2nd digit)':[ 2,  2,  2],
        'Level':              ['A', 'A', 'A'],
        'Cell':               [1, 2, 3],
        'ABB ZONE':           ['HRD', 'HRD', 'HRD'],
        'ABB LOCATION':       ['ABF', 'ABF', 'ABF'],
        'ABB FLOOR':          [1, 1, 1],
        'ABB RACK NO':        [2, 2, 2],
        'ABB LEVEL IN RACK':  ['C', 'D', 'B'],
        'ABB CELL':           [0, 0, 0],
        'ABB NO':             [1, 4, 5],
    }
    st.dataframe(pd.DataFrame(sample_data))

    st.markdown("""
**How model columns work (fully automatic):**

The tool scans your column headers and picks the **first 4 columns** that are not
Part No / Description / Qty / Location / ABB columns.  
Whatever text you put in those headers (e.g. `135 KW`, `60 KW`, `C`, `4W`) becomes
the label shown in the MTM box — **no code changes needed**.

Put the quantity for each vehicle type in the corresponding cell; leave blank if N/A.

ℹ️ Column names are case-insensitive for standard fields (Part No, Desc, Qty/Bin, etc.).
    """)

if __name__ == "__main__":
    main()
