import streamlit as st
import logging
import io
import re
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
from docx import Document
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def convert_pdf_to_images(pdf_file, poppler_path=None):
    """
    Converts a PDF file (provided as a BytesIO stream) into a list of PIL.Image objects.

    Args:
        pdf_file (BytesIO): The uploaded PDF file.
            Example: st.file_uploader("Upload PDF", type=["pdf"]) returns a file-like object.
        poppler_path (str, optional): Path to Poppler binaries if not in PATH.
            Example (Windows): r'C:\poppler\bin'. Defaults to None.

    Returns:
        list: A list of PIL.Image objects, each representing a page of the PDF.
            Example: [PIL.Image.Image, PIL.Image.Image, ...]
    """
    try:
        pdf_file.seek(0)
        pdf_bytes = pdf_file.read()
        images = convert_from_bytes(pdf_bytes, poppler_path=poppler_path)
        logging.info("Converted PDF to %d image(s).", len(images))
        return images
    except Exception as e:
        logging.error("Error converting PDF to images: %s", e)
        st.error("Failed to process PDF file. Ensure that Poppler is installed and in PATH.")
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
        try:
            text = pytesseract.image_to_string(image)
            texts.append(text)
            logging.info("Extracted text from page %d", idx)
        except Exception as e:
            logging.error("Error extracting text from page %d: %s", idx, e)
            texts.append("")  # Append an empty string if extraction fails
    return texts


def create_excel_file(texts):
    """
    Creates an Excel file from a list of extracted text strings after cleaning illegal characters.

    Args:
        texts (list): List of strings where each string is text extracted from a PDF page.
            Example: ["Page 1 text", "Page 2 text", ...]

    Returns:
        BytesIO: An in-memory bytes buffer containing the Excel file.
            Example: BytesIO object ready for download.
    """
    # Clean text by removing any illegal characters that Excel does not allow.
    def clean_text(text):
        return re.sub(ILLEGAL_CHARACTERS_RE, "", text) if text else text

    cleaned_texts = [clean_text(text) for text in texts]

    df = pd.DataFrame({
        'Page': list(range(1, len(cleaned_texts) + 1)),
        'Content': cleaned_texts
    })
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='ExtractedText')
        output.seek(0)
        logging.info("Excel file created with %d page(s).", len(cleaned_texts))
    except Exception as e:
        logging.error("Error creating Excel file: %s", e)
        st.error("Failed to create Excel file.")
    return output


def create_word_file(texts):
    """
    Creates a Word document from a list of extracted text strings.

    Args:
        texts (list): List of strings where each string is text extracted from a PDF page.
            Example: ["Page 1 text", "Page 2 text", ...]

    Returns:
        BytesIO: An in-memory bytes buffer containing the Word document.
            Example: BytesIO object ready for download.
    """
    document = Document()
    for idx, text in enumerate(texts, start=1):
        try:
            document.add_heading(f'Page {idx}', level=1)
            document.add_paragraph(text)
            logging.info("Added text for page %d to Word document.", idx)
        except Exception as e:
            logging.error("Error adding page %d to Word document: %s", idx, e)
    output = io.BytesIO()
    try:
        document.save(output)
        output.seek(0)
        logging.info("Word document created with %d page(s).", len(texts))
    except Exception as e:
        logging.error("Error saving Word document: %s", e)
        st.error("Failed to create Word document.")
    return output


# Streamlit User Interface
st.title("PDF to Excel/Word Converter")
st.write("Upload a scanned (non-readable) PDF and choose an export format.")

# File uploader accepts only PDF files.
uploaded_pdf = st.file_uploader("Upload a PDF file", type=["pdf"])

# Radio button to select export format.
export_format = st.radio("Select Export Format", ("Excel", "Word"))

# Optional: Specify the Poppler path if needed. For example, on Windows:
# poppler_path = r'C:\poppler\bin'
poppler_path = None  # Set to None if Poppler is already in PATH

if uploaded_pdf is not None:
    with st.spinner("Processing PDF..."):
        # Convert PDF pages to images.
        images = convert_pdf_to_images(uploaded_pdf, poppler_path=poppler_path)
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
            st.error("No images found in the PDF file. Check if the PDF is valid.")
