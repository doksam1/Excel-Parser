from openpyxl import load_workbook


def reverse_sheets_in_workbook(file_path, output_path=None):
    # Load the workbook
    wb = load_workbook(file_path)

    # Get the list of sheet names in reverse order
    reversed_sheet_names = wb.sheetnames[::-1]

    # Create a new workbook to store the reversed sheets
    new_wb = load_workbook(file_path)

    # Remove all sheets from the new workbook
    while new_wb.sheetnames:
        del new_wb[new_wb.sheetnames[0]]

    # Add sheets back in reversed order
    for sheet_name in reversed_sheet_names:
        source_sheet = wb[sheet_name]
        new_sheet = new_wb.create_sheet(title=sheet_name)

        for row in source_sheet.iter_rows():
            for cell in row:
                new_sheet[cell.coordinate].value = cell.value

    # Save the workbook to the output path
    if not output_path:
        output_path = file_path  # Overwrite the original file if no output path is specified
    new_wb.save(output_path)
    print(f"Workbook saved at {output_path}")


# Example usage
file_path = "your_workbook.xlsx"  # Replace with the path to your workbook
output_path = "reversed_workbook.xlsx"  # Specify a new file name or path
reverse_sheets_in_workbook(file_path, output_path)
