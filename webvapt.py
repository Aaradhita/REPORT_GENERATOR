from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import pandas as pd
from datetime import datetime
import os


# Load the Word template
template_path = r"c:\\Users\\Admin\\Downloads\\websampletemplate.docx"
if not os.path.exists(template_path):
    raise FileNotFoundError(f"Template not found at {template_path}")
template = DocxTemplate(template_path)

# Load the Excel tracker
excel_path = r"C:\\Users\\Admin\\Downloads\\webtrackersample.xlsx"
if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Excel tracker not found at {excel_path}")
df = pd.read_excel(excel_path)

# Validate required columns
required_columns = [
    'Affected Asset',
    'Observation / Vulnerability Title',
    'Severity',
    'Reference',
    'New or Repeat Observation',
    'Detailed observation / Vulnerable point',
    'Recommendations'
]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    raise ValueError(f"Missing columns in Excel file: {', '.join(missing_columns)}")

# Prepare vulnerability data
summary = []
details = []

for index, row in df.iterrows():
    folder_name = row.get("POC Folder", '').strip()
    poc_base_path= r"C:\\Users\\Admin\\Downloads\\poc"
    screenshots = []

    if folder_name:
        folder_path = os.path.join(poc_base_path, folder_name)
        if os.path.isdir(folder_path):
            for idx, fname in enumerate(sorted(os.listdir(folder_path)), start=1):
                if fname.lower().endswith(('.png', '.jpg', '.jpeg')):
                    img_path = os.path.join(folder_path, fname)
                    screenshots.append({
                    'label': f"Screenshot {idx}:",
                    'image': InlineImage(template, img_path, width=Mm(155))
                    })
        else:
            print(f"[!] Folder not found: {folder_path}")
    else:
        print(f"[!] Missing POC folder in row {index + 1}")
    summary.append({
        'sno': index + 1,
        'asset': row.get('Affected Asset', 'N/A'),
        'title': row.get('Observation / Vulnerability Title', 'N/A'),
        'cve': row.get('CVE/CWE', 'N/A'),
        'severity': row.get('Severity', 'N/A'),
        'reference': row.get('Reference', 'N/A'),
        'new_or_repeat': row.get('New or Repeat Observation', 'N/A'),
        
    })

    details.append({
        'title': row.get('Observation / Vulnerability Title', 'N/A'),
        'affected_assets': row.get('Affected Assets', 'N/A'),
        'description': row.get('Detailed observation / Vulnerable point', 'N/A'),
        'cve': row.get('CVE/CWE', 'N/A'),
        'recommendation': row.get('Recommendations', 'N/A'),
        'reference': row.get('Reference', 'N/A'),
        'status': row.get('New or Repeat Observation', 'N/A'),
        'screenshots': screenshots
        # Optionally add a screenshot:
        # 'screenshot': InlineImage(template, 'path_to_image.png', width=Mm(100))
    })

# Context passed to template
context = {
    'report_release_date': datetime.now().strftime("%d/%m/%Y"),
    'vulnerabilities_summary': summary,
    'vulnerabilities_details': details
}

# Check if summary is empty
if not summary:
    raise ValueError("No vulnerabilities found in the Excel file.")

# Render and save
output_path = r"C:\\Users\\Admin\\Downloads\\Final_Report.docx"
template.render(context)
template.save(output_path)

print(f"Report generated successfully: {output_path}")
