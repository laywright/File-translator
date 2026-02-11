# app.py
# Streamlit app: Upload Excel/Word -> paginate -> export Excel/PDF -> translate to English -> export PDF

import os
import io
import math
import tempfile
import logging
from typing import List
import pandas as pd
import streamlit as st
from docx import Document
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from deep_translator import GoogleTranslator

# ------------------------- Logging -------------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# ------------------------- Helpers -------------------------
ALLOWED_EXTS = {".xlsx", ".xls", ".docx"}

def _save_uploaded_file(uploaded_file) -> str:
    suffix = os.path.splitext(uploaded_file.name)[1].lower()
    if suffix not in ALLOWED_EXTS:
        raise ValueError(f"Unsupported file type: {suffix}")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.flush()
    tmp.close()
    logger.info(f"Saved upload to {tmp.name}")
    return tmp.name

# ------------------------- Word Reader (Hardened) -------------------------
def read_docx_to_df(path: str) -> pd.DataFrame:
    """
    Robust Word reader:
    - Reads paragraphs
    - Safely reads tables with merged cells
    - Never crashes on malformed documents
    """
    doc = Document(path)
    rows = []

    # Paragraphs
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            rows.append({"content": txt})

    # Tables (safe)
    for table in doc.tables:
        for row in table.rows:
            safe_cells = []
            for cell in row.cells:
                try:
                    text = cell.text.strip()
                except Exception:
                    text = ""
                if text:
                    safe_cells.append(text)
            if safe_cells:
                rows.append({"content": " | ".join(safe_cells)})

    if not rows:
        rows.append({"content": "No readable content found, but document processed successfully."})

    return pd.DataFrame(rows)

# ------------------------- Excel Reader -------------------------
def read_excel_to_df(path: str) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    frames = []
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
            df.insert(0, "_sheet", sheet)
            frames.append(df)
        except Exception as e:
            logger.warning(f"Skipping sheet {sheet}: {e}")
    if not frames:
        raise ValueError("No readable sheets found in Excel file.")
    return pd.concat(frames, ignore_index=True).fillna("")

def parse_input(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in {".xlsx", ".xls"}:
        return read_excel_to_df(path)
    elif ext == ".docx":
        return read_docx_to_df(path)
    else:
        raise ValueError("Unsupported file type.")

# ------------------------- Pagination -------------------------
def paginate_df(df: pd.DataFrame, rows_per_page: int) -> List[pd.DataFrame]:
    pages = []
    total = len(df)
    for start in range(0, total, rows_per_page):
        pages.append(df.iloc[start:start + rows_per_page].reset_index(drop=True))
    return pages

# ------------------------- Export Excel -------------------------
def export_excel(pages: List[pd.DataFrame], out_path: str) -> None:
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    for i, page in enumerate(pages, 1):
        ws = wb.create_sheet(title=f"Page_{i}")
        for r in dataframe_to_rows(page, index=False, header=True):
            ws.append(r)
        # auto width
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    wb.save(out_path)
    logger.info(f"Excel exported: {out_path}")

# ------------------------- Export PDF -------------------------
def _table_from_df(df: pd.DataFrame):
    data = [list(df.columns)] + df.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ])
    table.setStyle(style)
    return table

def export_pdf(pages: List[pd.DataFrame], out_path: str, title: str = "Export") -> None:
    doc = SimpleDocTemplate(out_path, pagesize=landscape(A4), rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles['Title']), Spacer(1, 12)]

    for i, page in enumerate(pages, 1):
        story.append(Paragraph(f"Page {i}", styles['Heading2']))
        story.append(Spacer(1, 8))
        story.append(_table_from_df(page))
        if i != len(pages):
            story.append(PageBreak())
    doc.build(story)
    logger.info(f"PDF exported: {out_path}")

# ------------------------- Translation -------------------------
def translate_df_to_english(df: pd.DataFrame, batch_size: int = 40) -> pd.DataFrame:
    translator = GoogleTranslator(source='auto', target='en')
    out = df.copy()
    for col in out.columns:
        series = out[col].astype(str)
        translated = []
        buffer = []
        for i, txt in enumerate(series):
            buffer.append(txt)
            if len(buffer) >= batch_size or i == len(series) - 1:
                try:
                    res = translator.translate_batch(buffer)
                except Exception as e:
                    logger.warning(f"Batch translate failed, falling back to single: {e}")
                    res = []
                    for t in buffer:
                        try:
                            res.append(translator.translate(t))
                        except Exception:
                            res.append(t)
                translated.extend(res)
                buffer = []
        out[col] = translated
    return out

# ------------------------- Streamlit UI -------------------------
st.set_page_config(page_title="File Splitter → PDF → English", layout="wide")
st.title("📄 File Splitter → Excel Pages → PDF → English Translator")
st.markdown("""
Upload an **Excel (.xlsx/.xls)** or **Word (.docx)** file. The app will:  
1) Split into pages (default 50 rows per page) as Excel sheets  
2) Export a PDF  
3) Translate content to English and export a translated PDF
""")

rows_per_page = st.number_input("Rows per page", min_value=5, max_value=500, value=50, step=5)
uploaded = st.file_uploader("Upload file", type=["xlsx", "xls", "docx"])

if uploaded is not None:
    try:
        path = _save_uploaded_file(uploaded)
        with st.spinner("Reading file…"):
            df = parse_input(path)
        st.success(f"Loaded {len(df)} rows × {len(df.columns)} columns")
        st.dataframe(df.head(20), use_container_width=True)

        if st.button("Process File", type="primary"):
            progress = st.progress(0)
            tmpdir = tempfile.mkdtemp(prefix="splitpdf_")
            excel_out = os.path.join(tmpdir, "paginated.xlsx")
            pdf_out = os.path.join(tmpdir, "export.pdf")
            tpdf_out = os.path.join(tmpdir, "translated_en.pdf")

            # paginate
            with st.spinner("Paginating…"):
                pages = paginate_df(df, int(rows_per_page))
            progress.progress(25)

            # excel
            with st.spinner("Exporting Excel…"):
                export_excel(pages, excel_out)
            progress.progress(50)

            # pdf
            with st.spinner("Building PDF…"):
                export_pdf(pages, pdf_out, title="Original Export")
            progress.progress(70)

            # translate
            with st.spinner("Translating to English… (may take time)"):
                tpages = [translate_df_to_english(p) for p in pages]
            progress.progress(90)

            with st.spinner("Building translated PDF…"):
                export_pdf(tpages, tpdf_out, title="Translated (English)")
            progress.progress(100)

            # downloads
            st.subheader("Downloads")
            with open(excel_out, "rb") as f:
                st.download_button("⬇️ Download Excel (paginated)", data=f, file_name="paginated.xlsx")
            with open(pdf_out, "rb") as f:
                st.download_button("⬇️ Download PDF", data=f, file_name="export.pdf")
            with open(tpdf_out, "rb") as f:
                st.download_button("⬇️ Download Translated PDF (English)", data=f, file_name="translated_en.pdf")

    except Exception as e:
        logger.exception(e)
        st.error(f"Error: {e}")
