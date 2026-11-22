"""PDF text extraction using pymupdf (primary), pdfplumber (fallback), and OCR (final fallback)."""

import os
import re

import fitz  # pymupdf

def _extract_with_pymupdf(file_path):
    """Extract text using pymupdf. Returns (text, metadata_dict)."""
    doc = fitz.open(file_path)
    text_parts = []
    for page in doc:
        text_parts.append(page.get_text())
    text = "\n\n".join(text_parts)

    meta = doc.metadata or {}
    metadata = {}
    for key in ("title", "author", "subject", "creator"):
        val = meta.get(key)
        if val:
            metadata[key] = str(val).strip()

    for date_key in ("creationDate", "modDate"):
        raw = meta.get(date_key, "")
        if raw:
            cleaned = str(raw).replace("D:", "").strip()
            metadata["pub_date"] = cleaned[:10].replace("'", "")
            break

    metadata["page_count"] = doc.page_count
    metadata["file_name"] = os.path.basename(file_path)
    doc.close()
    return text, metadata


def _extract_with_pdfplumber(file_path):
    """Fallback extraction using pdfplumber."""
    import pdfplumber
    text_parts = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)
    text = "\n\n".join(text_parts)
    metadata = {"file_name": os.path.basename(file_path), "method": "pdfplumber"}
    return text, metadata


def _extract_with_ocr(file_path):
    """OCR fallback: render PDF pages as images and run EasyOCR.
    Supports 80+ languages including Chinese, English, Japanese, Korean,
    French, German, Spanish, Russian, Italian, Portuguese, Arabic etc."""
    import easyocr
    import io

    reader = easyocr.Reader(
        ['ch_sim', 'en', 'ja', 'ko', 'fr', 'de', 'es', 'ru', 'it', 'pt', 'ar'],
        gpu=False
    )

    doc = fitz.open(file_path)
    text_parts = []

    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img_bytes = pix.tobytes("png")
        results = reader.readtext(img_bytes, detail=0)
        if results:
            text_parts.append("\n".join(results))

    doc.close()
    text = "\n\n".join(text_parts)
    metadata = {"file_name": os.path.basename(file_path), "method": "ocr"}
    return text, metadata


def _has_meaningful_text(text):
    """Check if text has meaningful content (not just whitespace or pure PUA garbage)."""
    if not text or not text.strip():
        return False
    meaningful = sum(1 for c in text if c.isalnum() or '一' <= c <= '鿿'
                     or '぀' <= c <= 'ヿ' or '가' <= c <= '힯'
                     or 'Ѐ' <= c <= 'ӿ' or '؀' <= c <= 'ۿ')
    return meaningful > 10


def extract_text_from_pdf(file_path):
    """Extract all text from a PDF file, with pdfplumber and OCR fallbacks."""
    text, _ = extract_text_and_metadata(file_path)
    return text


def extract_text_and_metadata(file_path):
    """Extract text and PDF metadata. Tries pymupdf → pdfplumber → OCR."""
    text, metadata = _extract_with_pymupdf(file_path)

    if not _has_meaningful_text(text):
        try:
            text2, meta2 = _extract_with_pdfplumber(file_path)
            if _has_meaningful_text(text2):
                text = text2
                metadata.update(meta2)
        except Exception:
            pass

    if not _has_meaningful_text(text):
        try:
            text3, meta3 = _extract_with_ocr(file_path)
            if _has_meaningful_text(text3):
                text = text3
                metadata.update(meta3)
        except ImportError:
            pass
        except Exception:
            pass

    return text, metadata
