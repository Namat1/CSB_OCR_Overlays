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

# Funktion: Overlays mit Namen hinzufügen
def add_overlays_with_text_on_top(pdf_file, page_name_map, name_x=200, name_y=750, extra_x=None, extra_y=None, name_color="#FF0000", extra_color="#0000FF"):
    reader = PdfReader(pdf_file)
    writer = PdfWriter()

    # Farbwerte in RGB umwandeln
    name_color_rgb = HexColor(name_color)
    extra_color_rgb = HexColor(extra_color)

    # Feste Texte
    text1 = "Name Fahrer: __________________________________"
    text3 = "Rolli Anzahl: ____________"
    text4 = "LKW: _____________"

    for page_number, page in enumerate(reader.pages):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)

        # Erstes Overlay (großes Rechteck)
        overlay_x, overlay_y, overlay_width, overlay_height = 177, 742, 393, 64
        can.setFillColorRGB(1, 1, 1)  # Weißer Hintergrund
        can.rect(overlay_x, overlay_y, overlay_width, overlay_height, fill=True, stroke=False)

        # Zweites Overlay (rechts oben)
        overlay2_x, overlay2_y, overlay2_width, overlay2_height = 425, 747, 202, 81
        can.setFillColorRGB(1, 1, 1)
        can.rect(overlay2_x, overlay2_y, overlay2_width, overlay2_height, fill=True, stroke=False)

        # Drittes Overlay (unten)
        overlay3_x, overlay3_y, overlay3_width, overlay3_height = 40, 640, 475, 20
        can.setFillColorRGB(1, 1, 1)
        can.rect(overlay3_x, overlay3_y, overlay3_width, overlay3_height, fill=True, stroke=False)

        # Feste Texte hinzufügen
        can.setFillColorRGB(0, 0, 0)
        can.setFont("Courier-Bold", 12)
        y_text_position = overlay_y + overlay_height - 20
        line_spacing = 30

        # Texte zeichnen
        can.drawString(overlay_x + 12, y_text_position, text1)
        can.drawString(overlay_x + 12, y_text_position - line_spacing, text3)
        can.drawString(overlay_x + 212, y_text_position - line_spacing, text4)

        # Warntext in Rot
        fixed_text = "!!! Achtung !!! Zwingend gesamtes Leergut abräumen."
        can.setFillColorRGB(1, 0, 0)
        can.setFont("Courier-Bold", 14)
        fixed_text_x, fixed_text_y = 75, 650
        can.drawString(fixed_text_x, fixed_text_y, fixed_text)

        # Unterstreichung
        text_width = can.stringWidth(fixed_text, "Courier-Bold", 14)
        underline_y = fixed_text_y - 5
        can.setLineWidth(1)
        can.line(fixed_text_x, underline_y, fixed_text_x + text_width, underline_y)

        # Namen und zusätzliche Werte
        if page_number in page_name_map:
            combined_name, extra_value = page_name_map[page_number]
            can.setFillColor(name_color_rgb)  # Dynamische Farbe für Namen
            can.setFont("Courier-Bold", 20)
            can.drawString(name_x, name_y, combined_name)

            if extra_value and extra_x is not None and extra_y is not None:
                can.setFillColor(extra_color_rgb)  # Dynamische Farbe für Extra-Wert
                can.setFont("Courier-Bold", 22)
                can.drawString(extra_x, extra_y, extra_value)

        # Canvas speichern und zurücksetzen
        can.save()
        packet.seek(0)

        # Prüfen, ob Daten vorhanden sind
        if packet.getbuffer().nbytes == 0:
            st.error(f"Fehler beim Generieren des Overlays auf Seite {page_number}.")
            continue

        # Overlay-Seite lesen und zusammenführen
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
            # Sichere Umwandlung in Strings
            name_spalte_4 = str(match.iloc[0]['Name_4']) if not pd.isna(match.iloc[0]['Name_4']) else ""
            name_spalte_7 = str(match.iloc[0]['Name_7']) if not pd.isna(match.iloc[0]['Name_7']) else ""
            name_spalte_12 = str(match.iloc[0]['Name_12']) if not pd.isna(match.iloc[0]['Name_12']) else ""

            # E- vor Spalte 12 setzen, falls vorhanden
            extra_value = f"E-{name_spalte_12}" if name_spalte_12 else ""

            # Namen kombinieren
            combined_name = ", ".join(filter(None, [name_spalte_4, name_spalte_7]))
            page_name_map[page_number] = (combined_name, extra_value)
    return page_name_map

# Streamlit App
st.title("CSB Tourenplan")

# PDF-Upload
uploaded_pdf = st.file_uploader("Lade eine PDF-Datei hoch", type=["pdf"])
uploaded_excel = st.file_uploader("Lade eine Excel-Tabelle hoch", type=["xlsx"])

# Color Picker für Namen und E-Wert
name_color = st.color_picker("Wählen Sie eine Farbe für den Namen", "#FF0000")  # Standard Rot
extra_color = st.color_picker("Wählen Sie eine Farbe für den LKW", "#0000FF")  # Standard Blau

if uploaded_pdf and uploaded_excel:
    # Zeige den "Ausführen"-Button, sobald beide Dateien hochgeladen wurden
    if st.button("Ausführen"):
        # Ladebalken starten
        with st.spinner("Verarbeitung läuft..."):
            progress_bar = st.progress(0)

            # OCR-Zahlen extrahieren
            rect = (94, 48, 140, 75)  # Pixelbereich (x0, y0, x1, y1)
            time.sleep(1)  # Simuliere Arbeit
            progress_bar.progress(25)

            page_numbers = extract_numbers_from_pdf(uploaded_pdf, rect)
            time.sleep(1)  # Simuliere Arbeit
            progress_bar.progress(50)

            # Excel-Tabelle einlesen und bereinigen
            excel_data = pd.read_excel(uploaded_excel, sheet_name="Touren", header=0)
            relevant_data = excel_data.iloc[:, [0, 3, 6, 11]].dropna(how='all')  # Spalten 1 (TOUR), 4, 7 und 12
            relevant_data.columns = ['TOUR', 'Name_4', 'Name_7', 'Name_12']
            relevant_data['TOUR'] = relevant_data['TOUR'].astype(str)

            # Abgleich durchführen
            page_name_map = match_numbers_with_excel(page_numbers, relevant_data)
            time.sleep(1)  # Simuliere Arbeit
            progress_bar.progress(75)

            # Overlays hinzufügen und Namen ins PDF schreiben
            output_pdf = add_overlays_with_text_on_top(
                uploaded_pdf, 
                page_name_map, 
                name_x=285, 
                name_y=785, 
                extra_x=430, 
                extra_y=755, 
                name_color=name_color,  # Übergabe der Name-Farbe
                extra_color=extra_color  # Übergabe der E-Wert-Farbe
            )
            time.sleep(1)  # Simuliere Arbeit
            progress_bar.progress(100)

        st.success("Verarbeitung abgeschlossen!")
        
        # Download-Button
        st.download_button(
            label="Bearbeitetes PDF herunterladen",
            data=output_pdf,
            file_name="output_with_overlays_and_names.pdf",
            mime="application/pdf"
        )
