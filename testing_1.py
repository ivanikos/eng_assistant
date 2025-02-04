import os
from datetime import datetime
import xlsxwriter


files_catalog = []

import os
from datetime import datetime

def list_files_with_metadata(folder_path, skip_folder="2022 BACK UP"):
    # Ensure the folder path exists
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"Folder not found: {folder_path}")

    # List to store file details and skipped paths
    file_data = []
    skipped_paths = []

    # Walk through the folder
    for root, dirs, files in os.walk(folder_path):
        # Skip the specified folder
        dirs[:] = [d for d in dirs if d != skip_folder]

        for file in files:
            file_path = os.path.join(root, file)
            try:
                # Check if the file exists before processing
                if not os.path.exists(file_path):
                    skipped_paths.append(file_path)
                    continue

                # Get the last modified time
                last_modified_time = os.path.getmtime(file_path)
                last_modified_date = datetime.fromtimestamp(last_modified_time).strftime('%Y-%m-%d %H:%M:%S')

                # Add file details to the list
                file_data.append({
                    "File Name": file,
                    "Last Modified": last_modified_date,
                    "File Path": file_path
                })

            except (FileNotFoundError, PermissionError, OSError) as e:
                # Log inaccessible file paths
                skipped_paths.append(file_path)

    return file_data, skipped_paths

# Example usage
folder_path = r"N:\Piping\Plant 3D\\"  # Replace with your folder path
try:
    files, skipped = list_files_with_metadata(folder_path, skip_folder="2022 BACK UP")
    print("Files in the folder:")
    for file in files:
        print(f"File: {file['File Name']}, Last Modified: {file['Last Modified']}, Path: {file['File Path']}")
        files_catalog.append([file['File Name'], file['Last Modified'], file['File Path']])

    if skipped:
        print("\nSkipped Paths:")
        for path in skipped:
            print(f"Skipped: {path}")

except FileNotFoundError as e:
    print(e)
except Exception as e:
    print(f"An unexpected error occurred: {e}")


workbook_summary = xlsxwriter.Workbook(f'C:\\Users\ivaign\OneDrive - United Conveyor Corp\Desktop\P3DAd\\file_list.xlsx')
ws = workbook_summary.add_worksheet("summary")

for i, (one, two, three) in enumerate(files_catalog):
    ws.write(f"A{i}", one)
    ws.write(f"B{i}", two)
    ws.write(f"C{i}", three)

workbook_summary.close()
print("success!")
