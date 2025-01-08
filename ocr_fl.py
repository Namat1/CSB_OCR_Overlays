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

# Funktion für das Hinzufügen von Overlays und Texten
def add_overlays_with_text_on_top(pdf_file, text1, text2, text3, text4, text_x_offset, additional_text=""):
    reader = PdfReader(pdf_file)
    writer = PdfWriter()

    for page in reader.pages:
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

        # Texte auf Overlays hinzufügen
        can.setFillColorRGB(0, 0, 0)  # Schwarzer Text
        can.setFont("Courier-Bold", 12)
        y_text_position = overlay_y + overlay_height - 20
        line_spacing = 30

        can.drawString(overlay_x + text_x_offset, y_text_position, text1)
        can.drawString(overlay_x + text_x_offset + 240, y_text_position, text2)
        can.drawString(overlay_x + text_x_offset, y_text_position - line_spacing, text3)
        can.drawString(overlay_x + text_x_offset + 200, y_text_position - line_spacing, text4)

        # Fester Text mit roter Schrift
        fixed_text = "!!! Achtung !!! Zwingend gesamtes Leergut abräumen."
        can.setFillColorRGB(1, 0, 0)  # Rot
        can.setFont("Courier-Bold", 14)
        fixed_text_x, fixed_text_y = 75, 650
        can.drawString(fixed_text_x, fixed_text_y, fixed_text)

        # Unterstreichung des festen Texts
        text_width = can.stringWidth(fixed_text, "Courier-Bold", 14)
        underline_y = fixed_text_y - 5
        can.setLineWidth(1)
        can.line(fixed_text_x, underline_y, fixed_text_x + text_width, underline_y)


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

# Funktion: OCR-Nummern mit Excel abgleichen
def match_numbers_with_excel(page_numbers, tour_data):
    page_name_map = {}
    for page_number, ocr_number in page_numbers.items():
        match = tour_data[tour_data['TOUR'] == ocr_number]
        if not match.empty:
            page_name_map[page_number] = match.iloc[0]['Name']
    return page_name_map

# Streamlit App
st.title("PDF OCR und Excel-Abgleich")

# PDF-Upload
uploaded_pdf = st.file_uploader("Lade eine PDF-Datei hoch", type=["pdf"])
uploaded_excel = st.file_uploader("Lade eine Excel-Tabelle hoch", type=["xlsx"])

if uploaded_pdf and uploaded_excel:
    # OCR-Nummern extrahieren
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

    # Namen ins PDF schreiben
    output_pdf = write_name_to_pdf_page(uploaded_pdf, page_name_map)

    # Download-Button
    st.download_button(
        label="Bearbeitetes PDF herunterladen",
        data=output_pdf,
        file_name="output_with_names.pdf",
        mime="application/pdf"
    )
