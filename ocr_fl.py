import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from io import BytesIO
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import re
import time

def add_overlays_with_text_on_top(pdf_file, page_name_map, name_x=200, name_y=750, extra_x=None, extra_y=None, name_color="#FF0000", extra_color="#0000FF"):
    reader = PdfReader(pdf_file)
    writer = PdfWriter()

    name_color_rgb = HexColor(name_color)
    extra_color_rgb = HexColor(extra_color)

    text1 = "Name Fahrer: __________________________________"
    text3_part1 = "Rolli Anzahl - "
    text3_part2 = "ca:"
    text3_part3 = " ________"
    text4 = "LKW: _____________"

    for page_number, page in enumerate(reader.pages):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)

        overlay_x, overlay_y, overlay_width, overlay_height = 177, 742, 393, 64
        can.setFillColorRGB(1, 1, 1)
        can.rect(overlay_x, overlay_y, overlay_width, overlay_height, fill=True, stroke=False)

        overlay2_x, overlay2_y, overlay2_width, overlay2_height = 425, 747, 202, 81
        can.setFillColorRGB(1, 1, 1)
        can.rect(overlay2_x, overlay2_y, overlay2_width, overlay2_height, fill=True, stroke=False)

        overlay3_x, overlay3_y, overlay3_width, overlay3_height = 40, 640, 475, 20
        can.setFillColorRGB(1, 1, 1)
        can.rect(overlay3_x, overlay3_y, overlay3_width, overlay3_height, fill=True, stroke=False)

        can.setFillColorRGB(0, 0, 0)
        can.setFont("Courier-Bold", 12)
        y_text_position = overlay_y + overlay_height - 20
        line_spacing = 30

        can.drawString(overlay_x + 12, y_text_position, text1)

        # Highlight 'ca:' with a yellow background
        can.drawString(overlay_x + 12, y_text_position - line_spacing, text3_part1)

        # Draw yellow rectangle for highlighting
        highlight_x = overlay_x + 12 + can.stringWidth(text3_part1, "Courier-Bold", 12)
        highlight_y = y_text_position - line_spacing - 2
        highlight_width = can.stringWidth(text3_part2, "Courier-Bold", 12)
        highlight_height = 14  # Adjust to font size

        can.setFillColorRGB(1, 1, 0)  # Yellow background
        can.rect(highlight_x, highlight_y, highlight_width, highlight_height, fill=True, stroke=False)

        # Draw the text 'ca:' in black on top of the highlight
        can.setFillColorRGB(0, 0, 0)
        can.drawString(highlight_x, y_text_position - line_spacing, text3_part2)

        can.drawString(
            overlay_x + 12 + can.stringWidth(text3_part1 + text3_part2, "Courier-Bold", 12),
            y_text_position - line_spacing,
            text3_part3
        )

        can.drawString(overlay_x + 212, y_text_position - line_spacing, text4)

        # Highlight fixed_text with a yellow background
        fixed_text = "!!! Achtung !!! Zwingend gesamtes Leergut abräumen."
        fixed_text_x, fixed_text_y = 75, 650
        can.setFont("Courier-Bold", 14)

        # Berechne Höhe und Breite des Textes
        fixed_text_width = can.stringWidth(fixed_text, "Courier-Bold", 14)
        fixed_text_height = 14  # Schriftgröße als Höhe

        # Zeichne das gelbe Highlight (zentriert mit der Schrift)
        highlight_x = fixed_text_x
        highlight_y = fixed_text_y - fixed_text_height + 5  # Leichte Anpassung für Zentrierung
        highlight_width = fixed_text_width
        highlight_height = fixed_text_height + 6  # Leichte Anpassung für visuelle Balance

        can.setFillColorRGB(1, 1, 0)  # Gelber Hintergrund
        can.rect(highlight_x, highlight_y, highlight_width, highlight_height, fill=True, stroke=False)

        # Zeichne den Text in Rot
        can.setFillColorRGB(1, 0, 0)  # Rote Schrift
        can.drawString(fixed_text_x, fixed_text_y, fixed_text)

        text_width = can.stringWidth(fixed_text, "Courier-Bold", 14)
        underline_y = fixed_text_y - 5
        can.setLineWidth(1)
        can.line(fixed_text_x, underline_y, fixed_text_x + text_width, underline_y)

        if page_number in page_name_map:
            combined_name, extra_value = page_name_map[page_number]
            can.setFillColor(name_color_rgb)
            can.setFont("Courier-Bold", 20)
            can.drawString(name_x, name_y, combined_name)

            if extra_value and extra_x is not None and extra_y is not None:
                can.setFillColor(extra_color_rgb)
                can.setFont("Courier-Bold", 22)
                can.drawString(extra_x, extra_y, extra_value)

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
            # Sichere Umwandlung in Strings und Überprüfung auf 0
            name_spalte_4 = str(match.iloc[0]['Name_4']) if not pd.isna(match.iloc[0]['Name_4']) and match.iloc[0]['Name_4'] != 0 else ""
            name_spalte_7 = str(match.iloc[0]['Name_7']) if not pd.isna(match.iloc[0]['Name_7']) and match.iloc[0]['Name_7'] != 0 else ""
            name_spalte_12 = str(match.iloc[0]['Name_12']) if not pd.isna(match.iloc[0]['Name_12']) and match.iloc[0]['Name_12'] != 0 else ""

            # E- vor Spalte 12 setzen, falls vorhanden
            extra_value = f"E-{name_spalte_12}" if name_spalte_12 else ""

            # Namen kombinieren
            combined_name = ", ".join(filter(None, [name_spalte_4, name_spalte_7]))
            page_name_map[page_number] = (combined_name, extra_value)
    return page_name_map

# Streamlit App
st.title("CSB Tourenplan schreiben")


# Farbauswahl-Optionen
color_options = {
    "Rot": "#FF0000",
    "Blau": "#0000FF",
    "Grün": "#00FF00",
    "Gelb": "#FFFF00",
    "Orange": "#FFA500",
    "Lila": "#800080",
    "Türkis": "#40E0D0",
    "Pink": "#FFC0CB",
    "Grau": "#808080",
    "Schwarz": "#000000"
}

# Farbauswahl für Namen
st.subheader("Wählen Sie eine Farbe für den Fahrer - Namen")
name_color_name = st.selectbox("Farbe für Namen", list(color_options.keys()))
name_color = color_options[name_color_name]

# Farbauswahl für LKW
st.subheader("Wählen Sie eine Farbe für das LKW Kennzeichen")
extra_color_name = st.selectbox("Farbe für den LKW", list(color_options.keys()), index=1)
extra_color = color_options[extra_color_name]

# PDF-Upload
uploaded_pdf = st.file_uploader("CSB Ladelisten PDF hochladen", type=["pdf"])
uploaded_excel = st.file_uploader("Excel Tourenplan der aktuellen Woche hochladen", type=["xlsx"])

if uploaded_pdf and uploaded_excel:
    if st.button("Ausführen"):
        with st.spinner("Verarbeitung läuft..."):
            rect = (94, 48, 140, 75)
            page_numbers = extract_numbers_from_pdf(uploaded_pdf, rect)
            excel_data = pd.read_excel(uploaded_excel, sheet_name="Touren", header=0)
            relevant_data = excel_data.iloc[:, [0, 3, 6, 11]].replace(0, pd.NA).dropna(how='all')
            relevant_data.columns = ['TOUR', 'Name_4', 'Name_7', 'Name_12']
            relevant_data['TOUR'] = relevant_data['TOUR'].astype(str)
            page_name_map = match_numbers_with_excel(page_numbers, relevant_data)
            output_pdf = add_overlays_with_text_on_top(
                uploaded_pdf, page_name_map, name_x=285, name_y=785, extra_x=430, extra_y=755, 
                name_color=name_color, extra_color=extra_color
            )
        st.success("Verarbeitung abgeschlossen!")
        st.download_button("Ladelisten PDF herunterladen", data=output_pdf, file_name="output_with_overlays.pdf", mime="application/pdf")
