#!/usr/bin/env python3

import os
import re
import pytesseract
import pandas as pd
from pdf2image import convert_from_path

# Ask user for the folder containing PDFs
pdf_directory = input("Enter the path to the folder containing PDF documents: ").strip()

if not os.path.exists(pdf_directory):
    print("Error: The specified folder does not exist.")
    exit()

# CSV output path - Desktop
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
csv_output_path = os.path.join(desktop_path, "Data.csv")

patterns = {
    "Date & Time": r"(\d{1,2}/\d{1,2}/\d{2,4},?\s\d{1,2}:\d{2}\s?[APap][Mm])",
    "DOD ID": r"DOD ID:\s*(\d+)",
    "Stock Number": r"STOCKNUMBER\s*([A-Za-z0-9]+)",
    "Item Description": r"(MacBook Pro A\d+)",
    "Serial Number": r"SN:\s*([A-Za-z0-9]+)",
    "Components": r"COMPONENTS:\s*(.*?)\n\n",
    "Replacement SNs": r"REPLACEMENT SNs\.:.*"
}

# Store extracted data
data_list = []

# Process each PDF file in the directory/folder
for filename in os.listdir(pdf_directory):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_directory, filename)

        try:
            # Convert PDF to images - for OCR
            images = convert_from_path(pdf_path)

            ocr_text = ""
            for image in images:
                ocr_text += pytesseract.image_to_string(image) + "\n\n"

            # Extract specific data
            extracted_data = {"Filename": filename}
            for key, pattern in patterns.items():
                matches = re.findall(pattern, ocr_text, re.DOTALL)
                extracted_data[key] = matches[0] if matches else "Not Found"

            data_list.append(extracted_data)

        except Exception as e:
            print(f"Error processing {filename}: {e}")

# Convert extracted data into a DataFrame and save to CSV
df = pd.DataFrame(data_list)
df.to_csv(csv_output_path, index=False, encoding="utf-8")

print(f"\nExtraction complete! Data saved to: {csv_output_path}")
