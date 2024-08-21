# Run this script with the command below;
# python kontek_erp_9_project.py "P:\KONTEK\ENGINEERING\ELECTRICAL\Application Development\ERP\9. Electrical BOM Parsing"

import json
import os
import openpyxl
import logging
import re

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_bom(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb.active
    parts = {}
    errors = []

    # Extract header row to determine column indices dynamically
    headers = {}
    header_row = sheet[1]
    for idx, cell in enumerate(header_row, 1):
        headers[cell.value] = idx

    # Define the expected columns, map them to headers if they exist
    try:
        catalog_col = headers['CATALOG']
        manufacturer_col = headers['MANUFACTURE']
        description_col = headers['DESCRIPTION']
        quantity_col = headers.get('QTY', None)  
        supplier_col = headers.get('SUPPLIER', None)
    except KeyError as e:
        errors.append(f"Missing expected header: {str(e)}")
        return parts, errors

    for row in range(2, sheet.max_row + 1):
        part_number = sheet.cell(row=row, column=catalog_col).value
        if not part_number:
            continue

        manufacturer = sheet.cell(row=row, column=manufacturer_col).value or 'Unknown'
        description = sheet.cell(row=row, column=description_col).value or 'No Description'

        try:
            # Only treat as valid quantity if it's a number; otherwise, log an error
            raw_quantity = sheet.cell(row=row, column=quantity_col).value
            if raw_quantity is not None and re.match(r'^\d+$', str(raw_quantity)):
                quantity = int(raw_quantity)
            else:
                raise ValueError("Invalid quantity")
        except (ValueError, TypeError) as e:
            quantity = 0
            errors.append(f"Invalid quantity for part {part_number} at row {row}: {str(e)}")

        supplier = sheet.cell(row=row, column=supplier_col).value if supplier_col else 'Unknown'

        parts[part_number] = {
            'manufacturer': manufacturer,
            'description': description,
            'quantity': quantity,
            'supplier': supplier
        }

    return parts, errors

def save_json(bom_parts, output_filename='parts.json'):
    with open(output_filename, 'w') as f:
        json.dump(bom_parts, f, indent=4)

def main():
    directory = "P:/KONTEK/ENGINEERING/ELECTRICAL/Application Development/ERP/9. Electrical BOM Parsing"
    all_bom_parts = {}
    all_errors = {}

    for filename in os.listdir(directory):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            filepath = os.path.join(directory, filename)
            logging.info(f"Processing file: {filename}")
            parts, errors = load_bom(filepath)
            all_bom_parts[filename] = parts
            if errors:
                all_errors[filename] = errors
                logging.warning(f"Errors encountered in file {filename}: {errors}")
    
    save_json(all_bom_parts, 'parts.json')
    save_json(all_errors, 'errors.json')

    logging.info("BOM parsing and JSON saving process completed.")

if __name__ == '__main__':
    main()
