# Flask-based Report Generator Web App with Login, Upload, and Download Features
from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, flash, jsonify
from werkzeug.utils import secure_filename
from flask_login import LoginManager, login_user, login_required, logout_user, UserMixin, current_user
import os, zipfile, shutil
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import pandas as pd
from datetime import datetime

# Configuration
app = Flask(__name__)
app.secret_key = 'your-secret-key'
login_manager = LoginManager()
login_manager.init_app(app)

# Paths
UPLOAD_FOLDER = 'uploads'
REPORT_FOLDER = 'generated_reports'
TEMPLATES = {
    'network': 'C:\\Users\\Admin\\Downloads\\samplereport.docx',
    'web': 'C:\\Users\\Admin\\Downloads\\websampletemplate.docx',
    'server': 'C:\\Users\\Admin\\Downloads\\server hardening.docx',
    'console': 'C:\\Users\\Admin\\Downloads\\firewall.docx',
    'OG': 'C:\\Users\\Admin\\Downloads\\ogcertain.docx'
}

HEADERS = {
    'network': ['Sno', 'Affected Asset', 'Observation / Vulnerability Title', 'Affected Assets', 'Detailed observation / Vulnerable point', 'CVE/CWE', 'Severity', 'Recommendations', 'Reference', 'New or Repeat Observation', 'POC Folder'],
    'web': ['S. No', 'Domain Name', 'Vulnerability Name', 'Vulnerable URL', 'Severity', 'Description', 'CVE/CWE', 'Fixing Recommendation', 'Referrence', 'POC'],
    'server': ['sno', 'DESCRIPTION', 'CATEGORY', 'ADVICES', 'SOLUTION', 'NAME', 'SEVERITY', 'HOSTNAME'],
    'console': ['Sno', 'Affected Asset', 'Observation / Vulnerability Title', 'Affected Assets', 'Detailed observation / Vulnerable point', 'CVE/CWE', 'Severity', 'Recommendations', 'Reference', 'New or Repeat Observation', 'POC Folder'],
    'OG': ['Sno', 'Affected Asset', 'Observation / Vulnerability Title', 'Affected Assets', 'Detailed observation / Vulnerable point', 'CVE/CWE', 'Severity', 'Recommendations', 'Reference', 'New or Repeat Observation', 'POC Folder', 'Recomendation']
}

ALLOWED_EXTENSIONS = {'xlsx', 'zip'}

# Users
users = {
    'sudo': {'password': 'technical', 'role': 'admin'},
    'sakthi': {'password': 'sakthi', 'role': 'user'},
    'aaradhita': {'password': 'aaradhita', 'role': 'user'}
}

class User(UserMixin):
    def __init__(self, username):
        self.id = username
        self.role = users[username]['role']

@login_manager.user_loader
def load_user(user_id):
    if user_id in users:
        return User(user_id)
    return None

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_zip(zip_path, extract_to):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username in users and users[username]['password'] == password:
            session.pop('_flashes', None)
            user = User(username)
            login_user(user)
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid credentials')
    return render_template('login.html')

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', role=current_user.role)

@app.route('/generate', methods=['GET', 'POST'])
@login_required
def generate():
    if request.method == 'POST':
        report_type = request.form['report_type'].lower()
        excel_file = request.files['excel']
        zip_file = request.files['poc_zip']
        file_name = request.form['filename']

        templates = {
            'network': 'C:\\Users\\Admin\\Downloads\\samplereport.docx',
            'web': 'C:\\Users\\Admin\\Downloads\\websampletemplate.docx',
            'server': 'C:\\Users\\Admin\\Downloads\\server hardening.docx',
            'console': 'C:\\Users\\Admin\\Downloads\\firewall.docx',
            'og': 'C:\\Users\\Admin\\Downloads\\ogcertain.docx'
        }

        expected_headers = {
            'network': ['Sno', 'Affected Asset', 'Observation / Vulnerability Title', 'Affected Assets',
                        'Detailed observation / Vulnerable point', 'CVE/CWE', 'Severity',
                        'Recommendations', 'Reference', 'New or Repeat Observation', 'POC Folder'],
            'web': ['S. No', 'Domain Name', 'Vulnerability Name', 'Vulnerable URL', 'Severity',
                    'Description', 'CVE/CWE', 'Fixing Recommendation', 'Referrence', 'POC'],
            'server': ['sno', 'DESCRIPTION', 'CATEGORY', 'ADVICES', 'SOLUTION', 'NAME', 'SEVERITY', 'HOSTNAME'],
            'console': ['Sno', 'Affected Asset', 'Observation / Vulnerability Title', 'Affected Assets',
                        'Detailed observation / Vulnerable point', 'CVE/CWE', 'Severity',
                        'Recommendations', 'Reference', 'New or Repeat Observation', 'POC Folder'],
            'og': ['Sno', 'Affected Asset', 'Observation / Vulnerability Title', 'Affected Assets',
                   'Detailed observation / Vulnerable point', 'CVE/CWE', 'Severity',
                   'Recommendations', 'Reference', 'New or Repeat Observation', 'POC Folder', 'Recomendation'],
        }

        if report_type not in templates:
            flash('Invalid report type selected.')
            return redirect(url_for('generate'))

        template_path = templates[report_type]
        expected = expected_headers[report_type]

        if not (excel_file and allowed_file(excel_file.filename)) or not (zip_file and allowed_file(zip_file.filename)):
            flash('Invalid file types. Please upload both .xlsx and .zip.')
            return redirect(url_for('generate'))

        work_dir = os.path.join(UPLOAD_FOLDER, file_name)
        os.makedirs(work_dir, exist_ok=True)

        excel_path = os.path.join(work_dir, secure_filename(excel_file.filename))
        zip_path = os.path.join(work_dir, secure_filename(zip_file.filename))
        excel_file.save(excel_path)
        zip_file.save(zip_path)

        extract_path = os.path.join(work_dir, 'poc')
        extract_zip(zip_path, extract_path)

        # ✅ Read Excel and check headers
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            flash(f"Error reading Excel file: {e}")
            return redirect(url_for('generate'))

        actual_headers = list(df.columns)
        if not all(header in actual_headers for header in expected):
            flash(f"Header mismatch! Expected headers for '{report_type}' report: {expected}")
            return redirect(url_for('generate'))

        # ✅ If headers are OK, proceed to render report (customize based on type if needed)
        doc = DocxTemplate(template_path)
        summary, details = [], []

        for index, row in df.iterrows():
            row_data = {col: str(row.get(col, 'N/A')).strip() if not pd.isna(row.get(col)) else 'N/A' for col in expected_headers}

            # Example handling: You may create different logic per report_type here
            summary.append({key: row_data.get(key, 'N/A') for key in expected})
            details.append({key: row_data.get(key, 'N/A') for key in expected})

        context = {
            'report_type': report_type.capitalize(),
            'report_release_date': datetime.now().strftime('%d/%m/%Y'),
            'vulnerabilities_summary': summary,
            'vulnerabilities_details': details
        }

        final_path = os.path.join(REPORT_FOLDER, f"{file_name}.docx")
        doc.render(context)
        doc.save(final_path)

        flash('Report generated successfully.')
        return redirect(url_for('reports'))

    return render_template('generate.html')


@app.route('/reports')
@login_required
def reports():
    query = request.args.get('q', '').lower()
    if os.path.isdir(REPORT_FOLDER):
        files = sorted(
            os.listdir(REPORT_FOLDER),
            key=lambda x: os.path.getmtime(os.path.join(REPORT_FOLDER, x)),
            reverse=True
        )
        filtered_files = []
        for file in files:
            if query in file.lower():
                file_path = os.path.join(REPORT_FOLDER, file)
                modified_time = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%d-%m-%Y %H:%M:%S')
                filtered_files.append({'name': file, 'time': modified_time})
    else:
        filtered_files = []

    return render_template('reports.html', files=filtered_files, query=query)

@app.route('/download/<filename>')
@login_required
def download(filename):
    return send_from_directory(REPORT_FOLDER, filename, as_attachment=True)

@app.route('/suggest')
@login_required
def suggest():
    query = request.args.get('q', '').lower()
    suggestions = []
    if os.path.isdir(REPORT_FOLDER):
        for file in os.listdir(REPORT_FOLDER):
            if query in file.lower():
                suggestions.append(file)
    return jsonify(suggestions)

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
