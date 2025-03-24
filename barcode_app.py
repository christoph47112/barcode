import streamlit as st
import pandas as pd
from io import BytesIO
from barcode import Code128
from barcode.writer import ImageWriter
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows
import tempfile
import os

st.title("📦 Artikelliste mit Code128-Barcodes")

uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Entferne unerwünschte Spalten, falls vorhanden
    spalten_zum_entfernen = ["MTART", "Abt.", "WGR", "WGR-Bezeichnung", "Wertart."]
    df = df.drop(columns=[s for s in spalten_zum_entfernen if s in df.columns])

    st.success("Datei erfolgreich verarbeitet. Vorschau:")
    st.dataframe(df.head())

    if st.button("🚀 Excel mit Barcodes erzeugen"):
        # Neue Arbeitsmappe erstellen
        wb = Workbook()
        ws = wb.active
        ws.title = "Artikelliste mit Barcode"

        # Tabelle einfügen
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Barcodes generieren und in Spalte K einfügen
        for i, art_nr in enumerate(df["Art-Nr"].astype(str), start=2):
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
                Code128(art_nr, writer=ImageWriter()).write(tmp_img)
                tmp_img.close()
                img = XLImage(tmp_img.name)
                img.width = 150
                img.height = 40
                ws.add_image(img, f"K{i}")

        ws.column_dimensions["K"].width = 25

        # Speichern in BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Excel-Datei mit Barcodes erstellt!")

        st.download_button(
            label="📥 Excel-Datei herunterladen",
            data=output,
            file_name="Artikelliste_mit_Barcodes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
