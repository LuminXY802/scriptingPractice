import os
import xml.etree.ElementTree as ET 
from openpyxl import Workbook

# Directory Path to XML files
directory_path = "/mnt/hdd/scriptingPractice/dmarcFilePulling/reports"

# Global Variables for work and worksheet for spreadsheet file
workbook = Workbook()
worksheet = workbook.active

for filename in os.listdir(directory_path):
    if filename.endswith('.xml'):
        file_path = os.path.join(directory_path, filename)
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            for metadata in root:
                metadata = root.find("report_metadata")
                for report_metadata in metadata:
                    org_name = metadata.find("org_name").text
                    email = metadata.find("email").text
                    report_id = metadata.find("report_id").text
                worksheet.append([org_name, email, report_id])
            for public in root:
                public = root.find("public_published")
                if public is not None:
                    for public_published in public:
                        domain = public_published.find("domain").text
                        adkim = public_published.find("adkim").text
                        aspf = public_published.find("aspf").text
                        p = public_published.find("p").text
                        sp = public_published.find("sp").text
                        pct = public_published.find("pct").text
                    worksheet.append([domain, adkim, aspf, p, sp, pct])
            for record in root:
                record = root.find("row")
                if record is not None:
                    for row in record:
                        source_ip = row.find("source_ip").text
                        count = row.find("count").text
                        disposition = row.find("disposition").text
                        dkim = row.find("dkim").text
                        spf = row.find("spf").text
                        header_from = row.find("header_from").text
                    worksheet.append([source_ip, count, disposition, dkim, spf, header_from])
            print(f"Processed: {file_path}")
        except ET.ParseError as e:
            print(f"Error parsing {file_path}: {e}")



workbook.save(filename="parsedData.xlsx")

