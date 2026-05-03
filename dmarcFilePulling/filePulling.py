import os
import xml.etree.ElementTree as ET 
from openpyxl import Workbook

# Directory Path to XML files
directory_path = "/mnt/hdd/scriptingPractice/dmarcFilePulling/reports"

# Global Variables for work and worksheet for spreadsheet file
workbook = Workbook()
workbook.save(filename="parsedData.xlsx")

for filename in os.listdir(directory_path):
    if filename.endswith('.xml'):
        file_path = os.path.join(directory_path, filename)
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
             # Process root element here
            print(f"Processed: {file_path}")
        except ET.ParseError as e:
            print(f"Error parsing {file_path}: {e}")



