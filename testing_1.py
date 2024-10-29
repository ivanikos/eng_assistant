import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import os

"""Draft for migration tool for Plant 3D specs"""

# Path to the original zip file
path_to_spec = r"C:\Users\ivaign\OneDrive - United Conveyor Corp\Documents\Python_Projects\reference files\SubSYS_L.pspx"
temp_zip_path = f"{os.getcwd()}\spec_temp.zip"  # Temporary file for the new zip

# Open the original zip file and create a new zip file
with zipfile.ZipFile(path_to_spec, 'r') as zip_ref, zipfile.ZipFile(temp_zip_path, 'w') as new_zip:
    # Iterate through all files in the original zip
    for item in zip_ref.infolist():
        # If the current file is CatalogReferences.xml, modify it
        if item.filename == 'editor/CatalogReferences.xml':
            # Extract the XML file into memory
            with zip_ref.open(item.filename) as xml_file:
                tree = ET.parse(xml_file)
                root = tree.getroot()
                for elem in root.iter():
                # Make modifications to the XML (example: modify tag)
                    if "pcat" in elem.text:
                        elem.text = elem.text.replace("2025", "TEST injection")

                # Save the modified XML into a BytesIO object
                modified_xml = BytesIO()
                tree.write(modified_xml, encoding='utf-8', xml_declaration=True)
                modified_xml.seek(0)

                # Write the modified XML back to the new zip
                new_zip.writestr(item.filename, modified_xml.getvalue())
        else:
            # For all other files, copy them as-is into the new zip
            new_zip.writestr(item.filename, zip_ref.read(item.filename))

# Replace the original zip with the modified zip
os.replace(temp_zip_path, path_to_spec)

print("XML file modified and saved back into the zip archive.")
