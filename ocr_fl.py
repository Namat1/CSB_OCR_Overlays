import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import re

# Funktion für das Hinzufügen von Overlays und Texten
def add_overlays_with_text_on_top(pdf_file, text1, text2, text3, text4, text_x_offset):
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

# Funktion für die OCR-Zahlenextraktion
def extract_numbers_from_pdf(pdf_file, rect, lang="eng"):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    all_numbers = []

    for page_number in range(len(doc)):
        page = doc.load_page(page_number)
        pix = page.get_pixmap(dpi=72)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        cropped_img = img.crop(rect)
        cropped_img = cropped_img.resize((cropped_img.width * 2, cropped_img.height * 2), Image.Resampling.LANCZOS)
        text = pytesseract.image_to_string(cropped_img, lang=lang)
        numbers = re.findall(r"\d+", text)
        all_numbers.append(f"Seite {page_number + 1}: {', '.join(numbers)}")
    
    doc.close()
    return all_numbers

# Streamlit-App
st.title("PDF-Bearbeitung: Overlays, OCR und Download")

# PDF-Upload
uploaded_pdf = st.file_uploader("Lade eine PDF-Datei hoch", type=["pdf"])

if uploaded_pdf is not None:
    # Overlay-Texte
    TEXT1, TEXT2, TEXT3, TEXT4 = "Name Fahrer: ___________________", "LKW: ______________", "Rolli Anzahl: ____________", "Gewaschen?: _____________"
    TEXT_X_OFFSET = 12

    if st.button("Overlay hinzufügen und Zahlen extrahieren"):
        with st.spinner("Füge Overlays hinzu..."):
            output_pdf = add_overlays_with_text_on_top(uploaded_pdf, TEXT1, TEXT2, TEXT3, TEXT4, TEXT_X_OFFSET)
        
        with st.spinner("Extrahiere Zahlen..."):
            rect = (94, 48, 140, 75)  # Pixelbereich (x0, y0, x1, y1)
            extracted_numbers = extract_numbers_from_pdf(BytesIO(output_pdf.read()), rect)

        st.success("Bearbeitung abgeschlossen!")

        # Ergebnisse anzeigen
        st.write("Extrahierte Zahlen:")
        for num in extracted_numbers:
            st.write(num)

        # Download-Button
        st.download_button(
            label="Bearbeitetes PDF herunterladen",
            data=output_pdf,
            file_name="output_with_overlays_and_numbers.pdf",
            mime="application/pdf"
        )
