import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import re

# Funktion: Overlays hinzufügen
def add_overlays_with_text_on_top(pdf_file, page_name_map, text_x_offset=12):
    reader = PdfReader(pdf_file)
    writer = PdfWriter()

    for page_number, page in enumerate(reader.pages):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)

        # Erstes Overlay
        overlay_x, overlay_y, overlay_width, overlay_height = 177, 742, 393, 64
        can.setFillColorRGB(1, 1, 1)  # Weißer Hintergrund
        can.rect(overlay_x, overlay_y, overlay_width, overlay_height, fill=True, stroke=False)

        # Zweites Overlay
        overlay2_x, overlay2_y, overlay2_width, overlay2_height = 425, 747, 202, 81
        can.setFillColorRGB(1, 1, 1)  # Weißer Hintergrund
        can.rect(overlay2_x, overlay2_y, overlay2_width, overlay2_height, fill=True, stroke=False)

        # Drittes Overlay
        overlay3_x, overlay3_y, overlay3_width, overlay3_height = 40, 640, 475, 20
        can.setFillColorRGB(1, 1, 1)  # Weißer Hintergrund
        can.rect(overlay3_x, overlay3_y, overlay3_width, overlay3_height, fill=True, stroke=False)

        # Text hinzufügen (falls Name vorhanden)
        if page_number in page_name_map:
            name = page_name_map[page_number]
            can.setFillColorRGB(0, 0, 0)  # Schwarzer Text
            can.setFont("Courier-Bold", 12)
            can.drawString(overlay_x + text_x_offset, overlay_y + 20, f"Name: {name}")

        can.save()
        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        overlay_page = overlay_pdf.pages[0]

        page.merge_page(overlay_page)
        writer.add_page(page)

    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output

# Funktion: OCR-Zahlen aus PDF extrahieren (nur vierstellige Nummern)
def extract_numbers_from_pdf(pdf_file, rect, lang="eng"):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    page_numbers = {}

    for page_number in range(len(doc)):
        page = doc.load_page(page_number)
        pix = page.get_pixmap(dpi=72)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        cropped_img = img.crop(rect)
        cropped_img = cropped_img.resize((cropped_img.width * 2, cropped_img.height * 2), Image.Resampling.LANCZOS)
        text = pytesseract.image_to_string(cropped_img, lang=lang)
        numbers = re.findall(r"\b\d{4}\b", text)  # Nur vierstellige Nummern
        if numbers:
            page_numbers[page_number] = numbers[0]  # Nehme die erste gefundene Nummer
    doc.close()
    return page_numbers

# Funktion: OCR-Nummern mit Excel abgleichen
def match_numbers_with_excel(page_numbers, excel_data):
    page_name_map = {}
    for page_number, ocr_number in page_numbers.items():
        match = excel_data[excel_data['TOUR'] == ocr_number]
        if not match.empty:
            page_name_map[page_number] = match.iloc[0]['Name']  # Name aus Spalte 4
    return page_name_map

# Streamlit App
st.title("PDF OCR und Excel-Abgleich mit Overlays")

# PDF-Upload
uploaded_pdf = st.file_uploader("Lade eine PDF-Datei hoch", type=["pdf"])
uploaded_excel = st.file_uploader("Lade eine Excel-Tabelle hoch", type=["xlsx"])

if uploaded_pdf and uploaded_excel:
    # OCR-Zahlen extrahieren (nur vierstellige Nummern)
    rect = (94, 48, 140, 75)  # Pixelbereich (x0, y0, x1, y1)
    page_numbers = extract_numbers_from_pdf(uploaded_pdf, rect)
    st.write("Extrahierte OCR-Nummern pro Seite:", page_numbers)

    # Excel-Tabelle einlesen und bereinigen
    excel_data = pd.read_excel(uploaded_excel, sheet_name="Touren", header=0)
    relevant_data = excel_data.iloc[:, [0, 3]].dropna()  # Spalten 1 (TOUR) und 4 (Name)
    relevant_data.columns = ['TOUR', 'Name']
    relevant_data['TOUR'] = relevant_data['TOUR'].astype(str)

    # Abgleich durchführen
    page_name_map = match_numbers_with_excel(page_numbers, relevant_data)
    st.write("Gefundene Namen pro Seite:", page_name_map)

    # Overlays hinzufügen und Namen ins PDF schreiben
    output_pdf = add_overlays_with_text_on_top(uploaded_pdf, page_name_map)

    # Download-Button
    st.download_button(
        label="Bearbeitetes PDF herunterladen",
        data=output_pdf,
        file_name="output_with_overlays_and_names.pdf",
        mime="application/pdf"
    )
