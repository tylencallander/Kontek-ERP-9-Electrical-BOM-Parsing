# Run this script with the command below;
#   python script.py "P:\KONTEK\ENGINEERING\ELECTRICAL\Application Development\ERP\9. Electrical BOM Parsing"

import json
import os
import openpyxl
import sys

def load_bom(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb.active
    parts = {}
    errors = []

    for row in range(2, sheet.max_row + 1):
        part_number = sheet['E' + str(row)].value
        if not part_number:
            continue
        
        manufacturer = sheet['F' + str(row)].value or 'Unknown'
        description = sheet['G' + str(row)].value or 'No Description'
        
        try:
            quantity = int(sheet['H' + str(row)].value or 0)  
        except ValueError:
            quantity = 0
            errors.append(f"Invalid quantity for part {part_number} at row {row}.")
        
        try:
            unit_price = float(sheet['I' + str(row)].value or 0.0)  
        except ValueError:
            unit_price = 0.0
            errors.append(f"Invalid unit price for part {part_number} at row {row}.")

        if part_number in parts:
            parts[part_number]['quantity'] += quantity
            parts[part_number]['unit_price'] = max(parts[part_number]['unit_price'], unit_price)
        else:
            parts[part_number] = {
                'manufacturer': manufacturer,
                'description': description,
                'quantity': quantity,
                'unit_price': unit_price
            }

    return parts, errors

def save_json(data, filename):
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

def main(directory):
    all_parts = {}
    all_errors = []

    for filename in os.listdir(directory):
        if filename.endswith(('.xlsx', '.xls')):
            filepath = os.path.join(directory, filename)
            parts, errors = load_bom(filepath)
            all_parts.update(parts)
            all_errors.extend(errors)

    save_json(all_parts, 'parts.json')
    if all_errors:
        save_json(all_errors, 'errors.json')

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <directory>")
        sys.exit(1)

    directory = sys.argv[1]
    main(directory)
