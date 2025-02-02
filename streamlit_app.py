import streamlit as st
import logging
import io
import re
import pandas as pd
import tempfile

from pdf2image import convert_from_bytes
import pytesseract
from pytesseract import Output
from docx import Document
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from PIL import ImageDraw  # For drawing bounding boxes

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def convert_pdf_to_images(pdf_file, poppler_path=None):
    """
    Converts a PDF file (provided as a BytesIO stream) into a list of image objects.
    
    Args:
        pdf_file (BytesIO): The uploaded PDF file.
        poppler_path (str, optional): Path to Poppler binaries if not in PATH.
    
    Returns:
        list: A list of PIL.Image objects (one per page).
    """
    try:
        pdf_file.seek(0)
        pdf_bytes = pdf_file.read()
        images = convert_from_bytes(pdf_bytes, poppler_path=poppler_path)
        logging.info("Converted PDF to %d image(s).", len(images))
        return images
    except Exception as e:
        logging.error("Error converting PDF to images: %s", e)
        st.error("Failed to process PDF file. Ensure Poppler is installed and in PATH.")
        return []


def extract_text_from_images(images):
    """
    Extracts text from a list of image objects using Tesseract OCR.
    
    Args:
        images (list): List of PIL.Image objects.
    
    Returns:
        list: A list of strings (extracted text for each image).
    """
    texts = []
    for idx, image in enumerate(images, start=1):
        try:
            text = pytesseract.image_to_string(image)
            texts.append(text)
            logging.info("Extracted text from page %d.", idx)
        except Exception as e:
            logging.error("Error extracting text from page %d: %s", idx, e)
            texts.append("")
    return texts


def create_excel_file(texts):
    """
    Creates an Excel file from a list of text strings after cleaning illegal characters.
    
    Args:
        texts (list): List of extracted text strings.
    
    Returns:
        BytesIO: In-memory Excel file ready for download.
    """
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
    Creates a Word document from a list of text strings.
    
    Args:
        texts (list): List of extracted text strings.
    
    Returns:
        BytesIO: In-memory Word document ready for download.
    """
    document = Document()
    for idx, text in enumerate(texts, start=1):
        try:
            document.add_heading(f'Page {idx}', level=1)
            document.add_paragraph(text)
            logging.info("Added page %d to Word document.", idx)
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


def extract_tables_from_images_with_tesseract(images):
    """
    Extracts table-like data from images using Tesseract's TSV output.
    For each image, the TSV output is converted into a pandas DataFrame.
    
    Args:
        images (list): List of PIL.Image objects.
    
    Returns:
        list: A list of pandas DataFrames, one per image (page).
    """
    tables = []
    for idx, image in enumerate(images, start=1):
        try:
            df = pytesseract.image_to_data(image, output_type=Output.DATAFRAME)
            # Filter out rows with empty text
            df = df[df['text'].notna() & (df['text'] != "")]
            logging.info("Extracted TSV data for page %d with %d rows.", idx, len(df))
            tables.append(df)
        except Exception as e:
            logging.error("Error extracting TSV data from page %d: %s", idx, e)
    return tables


def draw_bounding_boxes_on_image(image, conf_threshold=60):
    """
    Draws red bounding boxes on the image around detected text regions.
    
    Args:
        image (PIL.Image): The input image.
        conf_threshold (int): Confidence threshold for drawing boxes.
    
    Returns:
        PIL.Image: Image with red bounding boxes drawn.
    """
    draw = ImageDraw.Draw(image)
    data = pytesseract.image_to_data(image, output_type=Output.DICT)
    n_boxes = len(data['level'])
    for i in range(n_boxes):
        try:
            conf = int(data['conf'][i])
        except ValueError:
            conf = 0
        if conf > conf_threshold:
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            draw.rectangle([x, y, x + w, y + h], outline="red", width=2)
    return image


# --- Streamlit User Interface ---

st.title("PDF Converter & Table Extractor (Tesseract)")
st.write("Upload a scanned PDF and preview detected content with red bounding boxes before export.")

# File uploader accepts only PDF files.
uploaded_pdf = st.file_uploader("Upload a PDF file", type=["pdf"])

# Radio button to select export format.
export_format = st.radio("Select Export Format", ("Excel", "Word", "Tables (Tesseract)"))

# Optional: Specify Poppler path if needed (e.g., on Windows, set to r'C:\poppler\bin').
poppler_path = None

if uploaded_pdf is not None:
    with st.spinner("Processing PDF..."):
        # Convert PDF pages to images.
        images = convert_pdf_to_images(uploaded_pdf, poppler_path=poppler_path)
    
    if not images:
        st.error("No images found in the PDF file. Check if the PDF is valid.")
    else:
        # Show preview of detected content with red bounding boxes
        with st.expander("Preview Detected Content with Red Bounding Boxes"):
            for idx, img in enumerate(images, start=1):
                preview_img = draw_bounding_boxes_on_image(img.copy())
                st.image(preview_img, caption=f"Page {idx}", use_column_width=True)
        
        if export_format == "Tables (Tesseract)":
            with st.spinner("Extracting table data via Tesseract TSV output..."):
                table_dfs = extract_tables_from_images_with_tesseract(images)
            if table_dfs:
                st.success(f"Extracted table data from {len(table_dfs)} page(s).")
                for idx, table_df in enumerate(table_dfs, start=1):
                    st.write(f"### Page {idx} TSV Data")
                    st.dataframe(table_df)
                    csv_data = table_df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        label=f"Download Page {idx} Data as CSV",
                        data=csv_data,
                        file_name=f"page_{idx}_table_data.csv",
                        mime="text/csv"
                    )
            else:
                st.error("No table-like data could be extracted from the PDF using Tesseract.")
        else:
            with st.spinner("Extracting OCR text from PDF images..."):
                extracted_texts = extract_text_from_images(images)
            st.success("Text extraction complete.")
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
