import os
import pandas as pd


def merge_excel_files_to_sheets(folder_path, output_filename='merged_file.xlsx'):
    """
    Merge all Excel files in a folder into one Excel file with multiple sheets.

    Args:
        folder_path (str): Path to the folder containing Excel files
        output_filename (str): Name of the output merged file (default: 'merged_file.xlsx')

    Returns:
        None (creates a new Excel file in the same folder)
    """
    # Get all Excel files in the folder
    excel_files = [f for f in os.listdir(folder_path)
                   if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]

    if not excel_files:
        print("No Excel files found in the specified folder.")
        return

    # Create a Pandas Excel writer object
    output_path = os.path.join(folder_path, output_filename)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for file in excel_files:
            # Read each Excel file
            file_path = os.path.join(folder_path, file)

            try:
                # Read all sheets (if multiple sheets exist in one file)
                excel_data = pd.read_excel(file_path, sheet_name=None)

                for sheet_name, df in excel_data.items():
                    # Clean sheet name to avoid Excel restrictions
                    # Excel sheet name max length is 31
                    clean_sheet_name = sheet_name[:31]
                    clean_sheet_name = clean_sheet_name.replace(
                        ':', '').replace('\\', '').replace('/', '')

                    # Write to the merged file
                    df.to_excel(
                        writer, sheet_name=f"{os.path.splitext(file)[0]}_{clean_sheet_name}", index=False)

            except Exception as e:
                print(f"Error processing file {file}: {str(e)}")
                continue

    print(
        f"Successfully merged {len(excel_files)} files into {output_filename}")


# Example usage:
# # Correct way - point to the folder containing Excel files
merge_excel_files_to_sheets(r'C:\Users\NSR-PC\Downloads\excel_projects')
# merge_excel_files_to_sheets(
#     r'C:\Users\NSR-PC\Downloads\excel_projects\')
