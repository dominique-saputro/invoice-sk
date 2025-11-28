# streamlit_invoice_clean_layout.py
import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image, PageBreak, KeepTogether
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from typing import List, Dict
import streamlit.components.v1 as components
import tempfile

# ---------------------------
# Config: use provided logo path (will be transformed by environment)
# ---------------------------
logo_path = "assets/logo.jpeg"

# ---------------------------
# Terbilang (simple Indonesian converter)
# ---------------------------
def terbilang(n: int) -> str:
    units = ["", "satu", "dua", "tiga", "empat", "lima", "enam",
             "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"]
    def _to_words(x):
        x = int(x)
        if x < 12:
            return units[x]
        if x < 20:
            return _to_words(x - 10) + " belas"
        if x < 100:
            q, r = divmod(x, 10)
            return _to_words(q) + " puluh" + ("" if r == 0 else " " + _to_words(r))
        if x < 200:
            return "seratus" + ("" if x == 100 else " " + _to_words(x - 100))
        if x < 1000:
            q, r = divmod(x, 100)
            return _to_words(q) + " ratus" + ("" if r == 0 else " " + _to_words(r))
        if x < 2000:
            return "seribu" + ("" if x == 1000 else " " + _to_words(x - 1000))
        if x < 1000000:
            q, r = divmod(x, 1000)
            return _to_words(q) + " ribu" + ("" if r == 0 else " " + _to_words(r))
        if x < 1000000000:
            q, r = divmod(x, 1000000)
            return _to_words(q) + " juta" + ("" if r == 0 else " " + _to_words(r))
        q, r = divmod(x, 1000000000)
        return _to_words(q) + " milyar" + ("" if r == 0 else " " + _to_words(r))
    if n == 0:
        return "nol rupiah"
    if n < 0:
        return "minus " + terbilang(abs(n))
    return (_to_words(n) + " rupiah").title()

# ---------------------------
# Build a single combined PDF containing multiple invoices (clean/tight layout)
# ---------------------------

def build_combined_invoices_pdf(invoices: List[Dict], logo_path: str = None) -> bytes:
    """
    invoices: list of dicts: { 'header': {invoice_no, date, customer}, 'items': [ {...}, ... ] }
    returns: bytes of PDF (all invoices in one document)
    """
    buffer = BytesIO()
    # very small margins for tight layout
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            leftMargin=10*mm, rightMargin=10*mm,
                            topMargin=10*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    # compact styles
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=9, leading=10)
    small_bold = ParagraphStyle("small_bold", parent=styles["Normal"], fontSize=9, leading=10, spaceAfter=2, fontName="Helvetica-Bold")
    title = ParagraphStyle("title", parent=styles["Heading1"], alignment=1, fontSize=14, leading=16)
    center_small = ParagraphStyle("center_small", parent=styles["Normal"], fontSize=9, alignment=1,leading=10, spaceAfter=2, fontName="Helvetica-Bold")
    ter_style = ParagraphStyle("ter", parent=styles["Normal"], fontSize=9, leading=11, italic=True)
    small_right = ParagraphStyle("small", parent=styles["Normal"], fontSize=9, leading=10,alignment=2)
    small_right_bold = ParagraphStyle("small", parent=styles["Normal"], fontSize=9, leading=10,alignment=2,spaceAfter=2, fontName="Helvetica-Bold")

    story = []

    # prepare logo image if exists
    logo_img = None
    try:
        if logo_path:
            logo_img = Image(logo_path, width=15*mm, height=15*mm)
    except Exception:
        logo_img = None

    col_widths = [33*mm, 20*mm, 27*mm, 15*mm, 25*mm, 25*mm, 19*mm, 25*mm]  # similar to before, tuned

    for idx, inv in enumerate(invoices):
        header = inv["header"]
        items = inv["items"]

        # --- top header (logo, company name, contact) ---
        company = Paragraph("<b><br></br>PT. SETIA KAWAN MAKMUR SEJAHTERA</b>", ParagraphStyle("company", alignment=1, fontSize=13))
        address = Paragraph("Office and Mill :<br/>JL. JAYENG KUSUMA VII / 12,<br/>DESA TAPAN KECAMATAN KEDUNGWARU,<br/>TULUNGAGUNG 66229<br/>JAWA TIMUR, INDONESIA", small)
        contact = Paragraph("Telephone : +62.355.323190<br/>Fax : +62.355.323187", small_right)

        header_table = Table([[logo_img if logo_img else "", company,""]],
            colWidths=[20*mm, 100*mm, 20*mm]
        )
        header_table.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("ALIGN", (1,0), (1,0), "CENTER"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ]))
        story.append(header_table)
        story.append(Spacer(1, 15))
        
        header_table = Table(
            [
             [address, "" , contact]],
            colWidths=[50*mm, 85*mm, 50*mm]
        )
        header_table.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("ALIGN", (1,0), (1,0), "CENTER"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ]))
        story.append(header_table)
        story.append(Spacer(1, 4))

        # Title (Invoice / Kwitansi)
        story.append(Paragraph("Invoice", title))
        story.append(Paragraph("<b>Kwitansi</b>", title))
        story.append(Spacer(1, 6))

        # Customer + Invoice info row (tight)
        cust = Paragraph(f"<b>Customer :</b><br/>{header.get('customer','')}", small)
        inv_info = Paragraph(f"No : {header.get('invoice_no','')}<br/>Tanggal : {header.get('date','')}", small_right)
        ci = Table([[cust,"", inv_info]], colWidths=[50*mm, 85*mm, 50*mm])
        ci.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
        story.append(ci)
        story.append(Spacer(1, 6))

        # Items table header + rows
        tbl_data = [[
            Paragraph("<b>JENIS BARANG</b>", center_small),
            Paragraph("<b>TANGGAL</b>", center_small),
            Paragraph("<b>NOTA</b>", center_small),
            Paragraph("<b>Qty</b>", center_small),
            Paragraph("<b>DPP</b>", center_small),
            Paragraph("<b>PPN</b>", center_small),
            Paragraph("<b>PPH Ps.22</b>", center_small),
            Paragraph("<b>JUMLAH</b>", center_small),
        ]]

        # helper to format amounts
        def fmt_amount(x):
            try:
                if x == "" or x is None:
                    return ""
                return f"Rp. {int(float(x)):,}"
            except Exception:
                return str(x or "")
            
        def fint_amount(x):
            try:
                if x == "" or x is None:
                    return ""
                return f"{int(float(x)):,}"
            except Exception:
                return str(x or "")

        for it in items:
            row = [
                Paragraph(str(it.get("description","")), small),
                Paragraph(str(it.get("sj_date","") or it.get("date","")), small),
                Paragraph(str(it.get("sj_no","") or it.get("nota","")), small),
                Paragraph(fint_amount(it.get("qty","")), small_right),
                Paragraph(fmt_amount(it.get("dpp","")), small_right),
                Paragraph(fmt_amount(it.get("ppn","")), small_right),
                Paragraph(fmt_amount(it.get("pph","")), small_right),
                Paragraph(
                    fmt_amount(
                        int(float(it.get("dpp") or 0)) +
                        int(float(it.get("ppn") or 0)) +
                        int(float(it.get("pph") or 0))
                    ) if it.get("dpp") is not None else "",
                    small_right
                )
            ]
            tbl_data.append(row)

        table = Table(tbl_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#e8f2fb")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (3,1), (3,-1), "RIGHT"),
            ("ALIGN", (4,1), (-1,-1), "RIGHT"),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ]))
        story.append(table)
        story.append(Spacer(1, 6))

        # Summary on right (tight box)
        total_dpp = sum(int(float(it.get("dpp") or 0)) for it in items)
        total_ppn = sum(int(float(it.get("ppn") or 0)) for it in items)
        total_pph = sum(int(float(it.get("pph") or 0)) for it in items)
        total_all = total_dpp + total_ppn + total_pph

        if header.get('customer','') == 'SETIA KAWAN ABADI':
            fakturs = []
            bupots = []
            for it in items:
                fakturs.append(it.get('faktur_no',''))
                bupots.append(it.get('bupot_no',''))
            faktur_str = ", ".join(dict.fromkeys(str(x) for x in fakturs if x))
            bupot_str = ", ".join(dict.fromkeys(str(x) for x in bupots if x))
            summary_data = [
                [Paragraph(f"No. Faktur: {faktur_str}", small),Paragraph("DPP", small), Paragraph(": Rp. ", small), Paragraph(fint_amount(total_dpp), small_right)],
                [Paragraph(f"No. Bupot: {bupot_str}", small),Paragraph("PPN", small), Paragraph(": Rp. ", small), Paragraph(fint_amount(total_ppn), small_right)],
                ["",Paragraph("PPH (0,1%)", small), Paragraph(": Rp. ", small), Paragraph(fint_amount(total_pph), small_right)],
                ["",Paragraph("<b>TOTAL BAYAR</b>", small_bold), Paragraph(": Rp. ", small_bold), Paragraph(f"<b>{fint_amount(total_all)}</b>", small_right_bold)],
            ]
            sum_table = Table(summary_data, colWidths=[126*mm,28*mm, 10*mm, 20*mm])
        else:
            summary_data = [
                [Paragraph("DPP", small), Paragraph(": Rp. ", small), Paragraph(fint_amount(total_dpp), small_right)],
                [Paragraph("PPN", small), Paragraph(": Rp. ", small), Paragraph(fint_amount(total_ppn), small_right)],
                [Paragraph("PPH (0,1%)", small), Paragraph(": Rp. ", small), Paragraph(fint_amount(total_pph), small_right)],
                [Paragraph("<b>TOTAL BAYAR</b>", small_bold), Paragraph(": Rp. ", small_bold), Paragraph(f"<b>{fint_amount(total_all)}</b>", small_right_bold)],
            ]
            sum_table = Table(summary_data, colWidths=[28*mm, 10*mm, 20*mm])
        sum_table.setStyle(TableStyle([
            ("ALIGN", (2,0), (2,-1), "RIGHT"),
            ("LEFTPADDING", (0,0), (-1,-1), 2),
            ("RIGHTPADDING", (0,0), (-1,-1), 2),
            ("TOPPADDING", (0,0), (-1,-1), 2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ]))
        # place summary at the right by wrapping in a right-aligned table cell
        wrapper = Table([[sum_table]], colWidths=[sum(col_widths)])
        wrapper.setStyle(TableStyle([("ALIGN", (0,0), (0,0), "RIGHT")]))
        story.append(wrapper)
        story.append(Spacer(1, 6))

        # Terbilang box (smaller, italic)
        ter = Paragraph(terbilang(total_all), ter_style)
        ter_box = Table([[ter]], colWidths=[sum(col_widths)])
        ter_box.setStyle(TableStyle([
            ("BOX", (0,0), (-1,-1), 0.5, colors.black),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
            ("TOPPADDING", (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ]))
        story.append(Paragraph("<b>Terbilang :</b>", small))
        story.append(Spacer(1, 6))
        story.append(ter_box)
        story.append(Spacer(1, 15))

        # Bank info (compact)
        bank_text = Paragraph(
            "Untuk pembayaran, mohon ditransfer ke rekening bank kami dengan detail sebagai berikut :")
        story.append(bank_text)
        story.append(Spacer(1, 6))
        story.append(Paragraph("<b>BANK MANDIRI TULUNGAGUNG</b>"))
        story.append(Paragraph("<b>AN. : PT. SETIA KAWAN MAKMUR SEJAHTERA</b>"))
        story.append(Paragraph("<b>NO. REK : 1710099088000</b>"))
        story.append(Spacer(1, 15))

        # Signature area (tight)
        story.append(Paragraph("Hormat Kami,"))
        story.append(Spacer(1, 50))
        story.append(Paragraph("<b>( FANI CHRISYANTI, S.E., M.Ak )</b>"))
        story.append(Paragraph("<b>Manager Keuangan</b>"))

        # Add page break between invoices except last
        if idx < len(invoices) - 1:
            story.append(PageBreak())

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------------------
# Preview helper
# ---------------------------
def preview_pdf(pdf_bytes: bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        temp_url = f"file://{tmp.name}"

    components.iframe(temp_url, height=800, scrolling=True)

# ---------------------------
# Streamlit app
# ---------------------------
st.set_page_config(layout="wide")
st.title("Invoice SK")

excel_file = st.file_uploader("Upload Excel", type=["xlsx", "xls", "csv"])

# ----- LOAD DATA (your code stays the same) -----
if excel_file:
    if excel_file.name.lower().endswith(".csv"):
        df = pd.read_csv(excel_file)
    else:
        df = pd.read_excel(excel_file)
        
    df.columns = [c.strip() for c in df.columns]

    colmap = {
        "TANGGAL": "date",
        "NO. FAKTUR":"faktur_no",
        "NO. PPH":"bupot_no",
        "NO": "invoice_no",
        "NAMA": "customer",
        "BARANG": "description",
        "QTY": "qty",
        "DPP": "dpp",
        "PPN": "ppn",
        "PPH": "pph",
        "TGL SJN": "sj_date",
        "NO. SJN": "sj_no"
    }

    df_cols_upper = {c.upper(): c for c in df.columns}
    rename_map = {}
    for k, v in colmap.items():
        if k in df_cols_upper:
            rename_map[df_cols_upper[k]] = v
    df = df.rename(columns=rename_map)

    required = ["date","faktur_no","bupot_no","invoice_no","customer","description","qty","dpp","ppn","pph","sj_date","sj_no"]
    for col in required:
        if col not in df.columns:
            df[col] = ""

    # --- AG GRID SELECTION ---
    gb = GridOptionsBuilder.from_dataframe(df)

    # Enable row selection
    gb.configure_selection(
        selection_mode="multiple",
        use_checkbox=True,
        rowMultiSelectWithClick=True  # Enables shift-click & ctrl-click
    )

    # Add filters + sorting
    gb.configure_default_column(
        filter=True,
        sortable=True,
        resizable=True,
        wrapText=False,
        autoHeight=True,
    )

    grid_options = gb.build()

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        fit_columns_on_grid_load=True,
        enable_enterprise_modules=False,
        height=400,
        theme="streamlit",
    )

    # Extract selected rows as a DataFrame
    selected = pd.DataFrame(grid_response["selected_rows"])

    if selected.empty:
        st.info("Select rows.")
    else:
        st.success(f"{len(selected)} rows selected — building grouped invoices.")
        # group by customer, invoice_no, date
        grouped = selected.groupby(["customer", "invoice_no", "date"], sort=False)

        invoices = []
        for (customer, inv_no, inv_date), group in grouped:
            header = {"invoice_no": inv_no, "date": str(inv_date), "customer": customer}
            items = group[["description","qty","dpp","ppn","pph","sj_date","sj_no","faktur_no","bupot_no"]].to_dict(orient="records")
            invoices.append({"header": header, "items": items})

        final_pdf = build_combined_invoices_pdf(invoices, logo_path=logo_path)

        st.subheader("Preview")
        preview_pdf(final_pdf)

        st.download_button("Download combined PDF", data=final_pdf, file_name="invoices_combined.pdf", mime="application/pdf")
        st.frame("invoices_combined.pdf")