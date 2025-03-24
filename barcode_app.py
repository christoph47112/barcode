import streamlit as st
import pandas as pd
from io import BytesIO
from barcode import Code128
from barcode.writer import ImageWriter
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.graphics.barcode import code128 as rl_code128
from reportlab.lib.units import mm
import tempfile
import os

st.set_page_config(page_title="Barcodes erstellen", layout="wide")
st.title("📦 Artikelliste mit Code128-Barcodes")

def encode_code128(text):
    return chr(204) + text + chr(206)

# Datei-Upload
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch (.xlsx)", type=["xlsx"])

# Ausgabeformat wählen
output_option = st.radio("📤 Wähle Ausgabeformat:", [
    "Excel mit Barcode-Bild",
    "Excel mit Barcode-Text (für Code128-Schrift)",
    "PDF mit formatierten Barcodes"
])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Unerwünschte Spalten entfernen
    spalten_zum_entfernen = ["MTART", "Abt.", "WGR", "WGR-Bezeichnung", "Wertart."]
    df = df.drop(columns=[s for s in spalten_zum_entfernen if s in df.columns])

    st.success("✅ Datei geladen:")
    st.dataframe(df.head())

    if st.button("🚀 Datei erzeugen"):
        if output_option == "Excel mit Barcode-Bild":
            wb = Workbook()
            ws = wb.active
            ws.title = "Barcodes als Bild"
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    if r_idx == 1:
                        cell.font = Font(bold=True)

            for i, art_nr in enumerate(df["Art-Nr"].astype(str), start=2):
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
                    Code128(art_nr, writer=ImageWriter()).write(tmp_img, options={
                        "module_width": 0.25,
                        "module_height": 10,
                        "text_distance": 0,
                        "quiet_zone": 1
                    })
                    tmp_img.close()
                    img = XLImage(tmp_img.name)
                    img.width = 120
                    img.height = 30
                    ws.add_image(img, f"K{i}")
            ws.column_dimensions["K"].width = 22
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            st.download_button(
                label="📥 Excel mit Barcode-Bild herunterladen",
                data=output,
                file_name="Artikelliste_mit_Barcodebildern.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif output_option == "Excel mit Barcode-Text (für Code128-Schrift)":
            df["Barcode"] = df["Art-Nr"].astype(str).apply(encode_code128)
            wb = Workbook()
            ws = wb.active
            ws.title = "Barcodes als Text"
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    if r_idx == 1:
                        cell.font = Font(bold=True)
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            st.download_button(
                label="📥 Excel mit Barcode-Text herunterladen",
                data=output,
                file_name="Artikelliste_mit_Barcodetext.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif output_option == "PDF mit formatierten Barcodes":
            buffer = BytesIO()
            c = canvas.Canvas(buffer, pagesize=landscape(A4))
            width, height = landscape(A4)
            x_margin = 20 * mm
            y_margin = 20 * mm
            x = x_margin
            y = height - y_margin
            line_height = 20 * mm
            max_lines_per_page = int((height - 2 * y_margin) // line_height)
            line_count = 0

            for index, row in df.iterrows():
                if line_count >= max_lines_per_page:
                    c.showPage()
                    y = height - y_margin
                    line_count = 0
                art_nr = str(row["Art-Nr"])
                art_bez = row["Art-Bez"]
                barcode = rl_code128.Code128(art_nr, barHeight=15 * mm, barWidth=0.4)
                barcode.drawOn(c, x, y - 15 * mm)
                c.setFont("Helvetica", 10)
                c.drawString(x + 80 * mm, y - 5 * mm, f"{art_nr} | {art_bez}")
                y -= line_height
                line_count += 1

            c.save()
            buffer.seek(0)

            st.download_button(
                label="📥 PDF mit formatierten Barcodes herunterladen",
                data=buffer,
                file_name="Artikelliste_mit_Barcodes.pdf",
                mime="application/pdf"
            )
