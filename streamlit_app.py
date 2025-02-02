import streamlit as st
import logging
import io
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
from docx import Document

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def convert_pdf_to_images(pdf_file):
    """
    Converts a PDF file (provided as a BytesIO stream) into a list of PIL.Image objects.

    Args:
        pdf_file (BytesIO): The uploaded PDF file.
               Example: st.file_uploader("Upload PDF", type=["pdf"]) returns a file-like object.

    Returns:
        list: A list of PIL.Image objects, each representing a page of the PDF.
              Example: [PIL.Image.Image, PIL.Image.Image, ...]
    """
    try:
        pdf_file.seek(0)
        pdf_bytes = pdf_file.read()
        images = convert_from_bytes(pdf_bytes)
        logging.info(f"Converted PDF to {len(images)} image(s).")
        return images
    except Exception as e:
        logging.error("Error converting PDF to images: %s", e)
        st.error("Failed to process PDF file.")
        return []


def extract_text_from_images(images):
    """
    Extracts text from a list of PIL.Image objects using OCR (pytesseract).

    Args:
        images (list): A list of PIL.Image objects.
               Example: [PIL.Image.Image, PIL.Image.Image, ...]

    Returns:
        list: A list of strings containing the extracted text for each image.
              Example: ["Text from page 1", "Text from page 2", ...]
    """
    texts = []
    for idx, image in enumerate(images, start=1):
        text = pytesseract.image_to_string(image)
        texts.append(text)
        logging.info("Extracted text from page %d", idx)
    return texts


def create_excel_file(texts):
    """
    Creates an Excel file from the list of extracted text strings (one per page).

    Args:
        texts (list): List of strings where each string is text extracted from a PDF page.
               Example: ["Page 1 text", "Page 2 text", ...]

    Returns:
        BytesIO: An in-memory bytes buffer containing the Excel file.
               Example: BytesIO object ready for download.
    """
    df = pd.DataFrame({
        'Page': list(range(1, len(texts) + 1)),
        'Content': texts
    })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ExtractedText')
    output.seek(0)
    logging.info("Excel file created with %d page(s).", len(texts))
    return output


def create_word_file(texts):
    """
    Creates a Word document from the list of extracted text strings (one per page).

    Args:
        texts (list): List of strings where each string is text extracted from a PDF page.
               Example: ["Page 1 text", "Page 2 text", ...]

    Returns:
        BytesIO: An in-memory bytes buffer containing the Word document.
               Example: BytesIO object ready for download.
    """
    document = Document()
    for idx, text in enumerate(texts, start=1):
        document.add_heading(f'Page {idx}', level=1)
        document.add_paragraph(text)
        logging.info("Added page %d to Word document.", idx)
    output = io.BytesIO()
    document.save(output)
    output.seek(0)
    logging.info("Word document created with %d page(s).", len(texts))
    return output


# Streamlit User Interface
st.title("PDF to Excel/Word Converter")
st.write("Upload a scanned (non-readable) PDF and choose an export format.")

# File uploader accepts only PDF files.
uploaded_pdf = st.file_uploader("Upload a PDF file", type=["pdf"])

# Radio button to select export format.
export_format = st.radio("Select Export Format", ("Excel", "Word"))

if uploaded_pdf is not None:
    with st.spinner("Processing PDF..."):
        # Convert PDF pages to images.
        images = convert_pdf_to_images(uploaded_pdf)
        if images:
            # Extract text from images.
            extracted_texts = extract_text_from_images(images)
            st.success("Text extraction complete.")

            # Provide download option based on user selection.
            if export_format == "Excel":
                excel_file = create_excel_file(extracted_texts)
                st.download_button(
                    label="Download Excel File",
                    data=excel_file,
                    file_name="extracted_text.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif export_format == "Word":
                word_file = create_word_file(extracted_texts)
                st.download_button(
                    label="Download Word Document",
                    data=word_file,
                    file_name="extracted_text.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("No images to process from the PDF file.")
