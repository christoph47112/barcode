import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

st.set_page_config(page_title="Code128 Textbarcodes", layout="wide")
st.title("🔤 Artikelliste mit Code128-Barcodes als Text")

# Hilfsfunktion für Code128-Konvertierung (nur ASCII-Zeichen, keine Sonderzeichen)
def encode_code128(text):
    # Startzeichen B: ASCII 204, Endzeichen: 206
    return chr(204) + text + chr(206)

uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Entferne unerwünschte Spalten
    spalten_zum_entfernen = ["MTART", "Abt.", "WGR", "WGR-Bezeichnung", "Wertart."]
    df = df.drop(columns=[s for s in spalten_zum_entfernen if s in df.columns])

    # Neue Barcode-Spalte mit Code128-Zeichen als Text
    df["Barcode"] = df["Art-Nr"].astype(str).apply(encode_code128)

    st.success("✅ Datei verarbeitet. Vorschau:")
    st.dataframe(df.head())

    if st.button("📥 Excel-Datei erzeugen"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Artikelliste mit Text-Barcode"

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    cell.font = Font(bold=True)

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="⬇️ Excel herunterladen",
            data=output,
            file_name="Artikelliste_Code128_Text.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
