from openpyxl import load_workbook


def reverse_sheets_in_workbook(file_path, output_path=None):
    # Load the workbook
    print("loading workbook")
    wb = load_workbook(file_path)

    #reverse the order of sheets using move_sheet
    wb._sheets = wb._sheets[::-1]

    #save
    if not output_path:
        output_path = file_path
    wb.save(output_path)
    print(f"Workbook saved at {output_path}")


# Example usage
file_path = "Leasing Report for Apartment (Weekly report).xlsx"  # Replace with the path to your workbook
output_path = "reversed_workbook.xlsx"  # Specify a new file name or path
reverse_sheets_in_workbook(file_path, output_path)
