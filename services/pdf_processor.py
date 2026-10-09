import re
import pdfplumber
from pytesseract import image_to_string
from PIL import Image


def _find_plates(text: str) -> set:
    """Extract possible Brazilian license plate patterns from a given text.

    Supports classic format (AAA-1234), Mercosul format (ABC1D23) and common variations.
    """
    patterns = [
        r"\b[0-9]{3}[A-Z]{1}[0-9]{2}[A-Z]{2}\b",   # e.g., 123A45BC
        r"\b[0-9]{3}[A-Z]{3}[0-9]{1}\b",          # e.g., 123ABC4
        r"\b[A-Z]{3}[0-9]{4}\b",                    # e.g., ABC1234
        r"\b[A-Z]{3}[0-9][A-Z][0-9]{2}\b"          # Mercosul e.g., ABC1D23
    ]
    found = set()
    upper = text.upper()
    for pat in patterns:
        for match in re.findall(pat, upper):
            found.add(match)
    return found


def extrair_placas_pdf(pdf_path: str) -> list:
    """Extract all license plates from a PDF file using pdfplumber + OCR.

    The function reads each page, extracts native text, runs OCR on the page image
    (useful for scanned PDFs), and returns a list of unique plate strings.
    """
    plates = set()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Text extraction (if PDF has embedded text)
            txt = page.extract_text() or ""
            plates.update(_find_plates(txt))
            # OCR on the rendered page image
            pil_img = page.to_image(resolution=300).original
            ocr_text = image_to_string(pil_img, lang='por')
            plates.update(_find_plates(ocr_text))
    return list(plates)
