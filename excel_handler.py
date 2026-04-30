import openpyxl
import os

def fill_excel_template(data, template_path, output_path):
    """
    Fills the Excel template with extracted data while preserving formulas.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found at {template_path}")

    # Load workbook
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active # Assuming data goes to the active sheet

    # Mapping of JSON keys to Excel cells
    # Note: These should match the structure of assets/template.xlsx
    mapping = {
        "Consumer Name": "B2",
        "Consumer Number": "B3",
        "Billing Date": "B4",
        "Billing Period": "B5",
        "Units Consumed (kWh)": "B6",
        "Sanctioned Load (kW)": "B7",
        "Connected Load (kW)": "B8",
        "Tariff Category": "B9",
        "Total Bill Amount": "B10",
        "Due Date": "B11",
        "Meter Number": "B12"
    }

    for key, cell_coord in mapping.items():
        if key in data and data[key] is not None:
            value = data[key]
            
            # Clean up string values (strip spaces)
            if isinstance(value, str):
                value = value.strip()
                try:
                    # Convert to number if possible for Excel formulas
                    if '.' in value:
                        value = float(value)
                    else:
                        value = int(value)
                except ValueError:
                    pass
            
            ws[cell_coord] = value
        else:
            # Explicitly set missing values to "N/A" for clarity
            ws[cell_coord] = "N/A"

    # Save to output path
    wb.save(output_path)
    return output_path

if __name__ == "__main__":
    # Test block
    # test_data = {"Consumer Name": "John Doe", "Units Consumed (kWh)": 150}
    # fill_excel_template(test_data, "assets/template.xlsx", "output.xlsx")
    pass
