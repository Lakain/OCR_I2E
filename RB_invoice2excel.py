import pytesseract
import pandas as pd
from PIL import Image, ImageFilter, ImageEnhance
from pdf2image import convert_from_path
import os

TESSERACT_OCR_PATH = os.getenv('TESSERACT_OCR_PATH')
POPPLER_PATH = os.getenv('POPPLER_PATH')

def preprocess_image(image: Image):
    # Convert to grayscale
    image = image.convert('L')
    # Enhance contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)
    # Apply sharpening
    image = image.filter(ImageFilter.SHARPEN)
    return image

def find_column_location(data, name, conf_threshold):
    filtered_data = data[data.conf > conf_threshold]
    column_loc = filtered_data[filtered_data['text'].str.contains(name, case=False)]
    return column_loc

def extract_column(data, name, conf_threshold, offset=0, width=0):
    column_loc = find_column_location(data, name, conf_threshold)
    column = data[(data.conf > conf_threshold) & (data.left >= column_loc.left.values[0]+offset) & (data.left < column_loc.left.values[0]+width)& (data.top >= column_loc.top.values[0])]
    return column

def join_words_in_line(data, line_num):
    desc = ''
    for text in data[data.line_num == line_num].text:
        if desc == '':
            desc = text
        else:
            desc = " ".join([desc, text])
    return desc

images = convert_from_path("test_data\S247247 RB Collection.pdf", dpi=300, poppler_path=POPPLER_PATH)

all_processed_image_data = []

for image in images:
    image_processed = preprocess_image(image)
    processed_image_data = pytesseract.image_to_data(image_processed, output_type=pytesseract.Output.DATAFRAME)
    all_processed_image_data.append(processed_image_data)

