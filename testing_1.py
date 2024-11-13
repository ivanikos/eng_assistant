import win32com.client
import os

dwg_file = r"C:\Users\ivaign\OneDrive - United Conveyor Corp\Documents\Python_Projects\Testing Area\C-54880-13-(TESTING).dwg"

#
#
# def read_dwg_with_objectdbx(file_path):
#     try:
#         # Create an ObjectDBX application object
#         dbx = win32com.client.Dispatch("ObjectDBX.AxDbDocument.26")  # Version may vary
#
#         # Open the DWG file in the background
#         dbx.Open(file_path)
#
#         # Access blocks in the drawing
#         for block in dbx.Blocks:
#             print(f"Block Name: {block.Name}")
#
#         # Access layers in the drawing
#         for layer in dbx.Layers:
#             print(f"Layer Name: {layer.Name}")
#
#         # Close the document after processing
#         dbx.Close()
#
#     except Exception as e:
#         print(f"Error accessing {file_path}: {e}")
#
# # Example usage
# read_dwg_with_objectdbx(dwg_file)

try:
    obj_dbx = win32com.client.Dispatch("ObjectDBX.AxDbDocument.25")  # Adjust "25" for version 2025
except Exception as e:
    print(f"Error initializing ObjectDBX: {e}")
