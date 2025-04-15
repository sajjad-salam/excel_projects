import os
import win32com.client as win32


def merge_excel_files_to_sheets():
    # Prompt user to select a folder
    folder_path = input(
        "Enter the folder path containing Excel files: ").strip()

    # Validate folder path
    if not os.path.isdir(folder_path):
        print("The folder does not exist!")
        return

    # Create a new Excel application instance
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = False  # Run in the background
    excel_app.DisplayAlerts = False  # Suppress alerts

    # Create a new workbook for output
    wb_dest = excel_app.Workbooks.Add()
    # Remove this sheet later if no data is added
    default_sheet = wb_dest.Sheets(1)

    # Initialize variables
    file_count = 0

    try:
        # Loop through all files in the folder
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)

            # Check if the file is an Excel file (xlsx, xls, xlsm) and not temporary
            if (file_name.endswith(".xlsx") or file_name.endswith(".xls") or file_name.endswith(".xlsm")) \
                    and not file_name.startswith("~$"):
                try:
                    # Open the source workbook
                    wb_source = excel_app.Workbooks.Open(
                        file_path, ReadOnly=True)
                    file_count += 1

                    # Copy each sheet from the source workbook to the destination workbook
                    for sheet in wb_source.Sheets:
                        # Generate a valid sheet name (max 31 chars, no invalid characters)
                        # Remove file extension
                        sheet_name = file_name[:file_name.rfind(".")]
                        sheet_name = clean_sheet_name(sheet_name)

                        # Ensure unique sheet names by appending a counter if necessary
                        original_sheet_name = sheet_name
                        counter = 1
                        while sheet_name in [ws.Name for ws in wb_dest.Sheets]:
                            sheet_name = f"{original_sheet_name}_{counter}"
                            counter += 1

                        # Copy the sheet to the destination workbook
                        sheet.Copy(After=wb_dest.Sheets(wb_dest.Sheets.Count))
                        wb_dest.Sheets(wb_dest.Sheets.Count).Name = sheet_name

                    # Close the source workbook without saving
                    wb_source.Close(SaveChanges=False)

                except Exception as e:
                    print(f"Error processing file '{file_name}': {e}")

        # Remove the default sheet if it's empty
        if default_sheet and wb_dest.Sheets.Count > 1:
            default_sheet.Delete()

        # Save the merged workbook
        output_file = os.path.join(folder_path, "Merged_Workbook.xlsx")
        # FileFormat=51 corresponds to .xlsx
        wb_dest.SaveAs(output_file, FileFormat=51)
        print(
            f"Successfully merged {file_count} files into 'Merged_Workbook.xlsx'.")

    finally:
        # Quit the Excel application
        excel_app.DisplayAlerts = True
        excel_app.Quit()


def clean_sheet_name(sheet_name):
    """Remove invalid characters from sheet names."""
    invalid_chars = r'\/?*[]:'
    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, "")
    return sheet_name[:31]  # Limit to 31 characters


# Run the function
merge_excel_files_to_sheets()
