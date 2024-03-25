import pytesseract
from PIL import Image
import pandas as pd
from pdf2image import convert_from_path
import os

TESSERACT_OCR_PATH = os.getenv('TESSERACT_OCR_PATH')
POPPLER_PATH = os.getenv('POPPLER_PATH')

# Ensure the path to the Tesseract binary is set in your environment variables or specify it here
# pytesseract.pytesseract.tesseract_cmd = TESSERACT_OCR_PATH

# Function to convert PDF to images
def pdf_to_images(pdf_path):
    return convert_from_path(pdf_path, poppler_path=POPPLER_PATH)

# Function to apply OCR on an image and return text
def image_to_text(image):
    return pytesseract.image_to_string(image, lang="eng")

# Function to extract invoice data from OCR'd text (customize this part as per your invoice structure)
def extract_invoice_data(text):
    # Example: Extract and return dummy data; adjust extraction logic based on your invoice layout
    data = {
        "Invoice Number": "123456",
        "Date": "2023-01-01",
        "Total": "$1000"
    }
    return data

# Main function to convert PDF to Excel
def pdf_to_excel(pdf_path, excel_path):
    images = pdf_to_images(pdf_path)
    all_data = []
    
    for image in images:
        text = image_to_text(image)
        invoice_data = extract_invoice_data(text)
        all_data.append(invoice_data)
    
    # Convert list of data to pandas DataFrame
    df = pd.DataFrame(all_data)
    
    # Save DataFrame to Excel file
    df.to_excel(excel_path, index=False)

# Example usage
pdf_path = "C:/Users/amasc/Documents/GitHub/OCR_I2E/test_data/Amekor_Invoice_01.pdf"
excel_path = 'output_invoice_data.xlsx'
pdf_to_excel(pdf_path, excel_path)
