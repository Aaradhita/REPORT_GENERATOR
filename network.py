from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import pandas as pd
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Load the Word template
template_path = r"C:\\Users\\Admin\\Downloads\\samplereport.docx"
if not os.path.exists(template_path):
    raise FileNotFoundError(f"Template not found at {template_path}")
template = DocxTemplate(template_path)

def format_hosts(ip_str):
    return " | ".join([i.strip() for i in str(ip_str).split(";") if i.strip()])

def generate_report(master_file, client_file, output_path):
    try:
        if master_file.endswith(".csv"):
            master_df = pd.read_csv(master_file)
        else:
            master_df = pd.read_excel(master_file)

        client_df = pd.read_csv(client_file)
        merged = pd.merge(client_df, master_df, on="Observation / Vulnerability Title", how="left")

        # Rename client "IP(s)" column to 'Affected Assets' (IP or Host)
        merged["Affected Assets"] = merged["IP(s)"].apply(format_hosts)
        merged.drop(columns=["IP(s)"], inplace=True)

        # Reorder columns including both Affected Asset (type from master) and Affected Assets (IP from client)
        final = merged[[  
            "Affected Asset", 
            "Observation / Vulnerability Title",
            "Affected Assets",
            "Detailed observation / Vulnerable point",
            "CVE/CWE",
            "Severity",
            "Recommendations",
            "Reference",
            "New or Repeat Observation"
        ]]

        final.insert(0, "S.NO", range(1, len(final) + 1))

        final.to_excel(output_path, index=False)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate report:\n{e}")
        return False

def generate_single_client():
    master_file = filedialog.askopenfilename(title="Select Master Excel Tracker", filetypes=[("Excel Files", "*.xlsx *.xls")])
    client_file = filedialog.askopenfilename(title="Select Single Client CSV", filetypes=[("CSV Files", "*.csv")])
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Report As")

    if master_file and client_file and save_path:
        if generate_report(master_file, client_file, save_path):
            messagebox.showinfo("Success", f"Report saved at:\n{save_path}")

def generate_multiple_clients():
    master_file = filedialog.askopenfilename(title="Select Master Excel Tracker", filetypes=[("Excel Files", "*.xlsx *.xls")])
    client_folder = filedialog.askdirectory(title="Select Client CSV Folder")

    if master_file and client_folder:
        out_folder = os.path.join(client_folder, "exported_reports")
        os.makedirs(out_folder, exist_ok=True)

        for filename in os.listdir(client_folder):
            if filename.endswith(".csv"):
                client_path = os.path.join(client_folder, filename)
                client_name = os.path.splitext(filename)[0]
                out_path = os.path.join(out_folder, f"{client_name}_report.xlsx")
                generate_report(master_file, client_path, out_path)

        messagebox.showinfo("Success", f"All client reports saved in:\n{out_folder}")

# GUI
window = tk.Tk()
window.title("Vulnerability Tracker Generator")
window.geometry("450x250")

label = tk.Label(window, text="Select Report Generation Mode", font=("Arial", 12))
label.pack(pady=20)

btn1 = tk.Button(window, text="Generate for Single Client", command=generate_single_client, bg="#d0f0c0", font=("Arial", 10))
btn1.pack(pady=10)

btn2 = tk.Button(window, text="Generate for Multiple Clients", command=generate_multiple_clients, bg="#add8e6", font=("Arial", 10))
btn2.pack(pady=10)

window.mainloop()

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


    # Build image list from folder
    if folder_name:
        folder_path = os.path.join(poc_base_path, folder_name)
        if os.path.isdir(folder_path):
            for fname in sorted(os.listdir(folder_path)):
                if fname.lower().endswith(('.png', '.jpg', '.jpeg')):
                    img_path = os.path.join(folder_path, fname)
                    screenshots.append(InlineImage(template, img_path, width=Mm(100)))
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
        'affected_assets': row.get('Affected Asset', 'N/A'),
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
