import logging
import os
import zipfile
from datetime import datetime

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_from_directory,
    session,
    url_for,
)
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from werkzeug.utils import secure_filename

from config import (
    ALLOWED_EXTENSIONS,
    DOC_TEMPLATE_DIR,
    REMEMBER_COOKIE_DURATION,
    REPORT_FOLDER,
    SECRET_KEY,
    SESSION_PERMANENT,
    TEMPLATE_METADATA,
    UPLOAD_FOLDER,
)
from services.report_generator import (  # noqa: E402
    ReportGenerationError,
    generate_report,
    normalize_report_type,
)
from services.template_loader import TemplateLoaderError
from utils.file_handler import (
    ensure_directory,
    extract_zip,
    is_allowed_file,
    save_uploaded_file,
)

logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config["SESSION_PERMANENT"] = SESSION_PERMANENT
app.config["REMEMBER_COOKIE_DURATION"] = REMEMBER_COOKIE_DURATION

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"

users = {
    "sudo": {"password": "technical", "role": "admin"}
}


class User(UserMixin):
    def __init__(self, username: str):
        self.id = username
        self.role = users.get(username, {}).get("role", "user")


@login_manager.user_loader
def load_user(user_id: str):
    if user_id in users:
        return User(user_id)
    return None


ensure_directory(DOC_TEMPLATE_DIR)
ensure_directory(UPLOAD_FOLDER)
ensure_directory(REPORT_FOLDER)


@app.route("/", methods=["GET"])
def index():
    logged_in = current_user.is_authenticated if hasattr(current_user, "is_authenticated") else False
    return render_template(
        "index.html",
        templates=list(TEMPLATE_METADATA.keys()),
        logged_in=logged_in,
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if username in users and users[username]["password"] == password:
            user = User(username)
            login_user(user, remember=False)
            session.permanent = False
            return redirect(url_for("dashboard"))
        flash("Invalid credentials")
    return render_template("login.html")


@app.route("/dashboard")
@login_required
def dashboard():
    return render_template("index.html", role=current_user.role)


@app.route("/generate", methods=["GET", "POST"])
@login_required
def generate():
    if request.method == "POST":
        report_type_raw = request.form.get("report_type", "").strip()
        report_type = normalize_report_type(report_type_raw)
        if report_type not in TEMPLATE_METADATA:
            flash(f"Unsupported report type: '{report_type_raw}'.")
            return redirect(url_for("generate"))

        excel_file = request.files.get("excel")
        zip_file = request.files.get("poc_zip")
        file_name = (request.form.get("filename", "") or "").strip() or f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        if not excel_file or not is_allowed_file(excel_file.filename, ALLOWED_EXTENSIONS) or not excel_file.filename.lower().endswith(".xlsx"):
            flash("Invalid Excel. Please upload a .xlsx file.")
            return redirect(url_for("generate"))

        uses_screenshots = TEMPLATE_METADATA[report_type]["uses_screenshots"]
        if uses_screenshots:
            if not zip_file or not is_allowed_file(zip_file.filename, ALLOWED_EXTENSIONS) or not zip_file.filename.lower().endswith(".zip"):
                flash("Invalid PoC ZIP. Please upload a .zip for this template.")
                return redirect(url_for("generate"))

        work_dir = ensure_directory(os.path.join(UPLOAD_FOLDER, secure_filename(file_name)))

        try:
            excel_path = save_uploaded_file(excel_file, work_dir)
        except Exception as exc:
            logger.exception("Excel upload failed")
            flash(f"Unable to save Excel file: {exc}")
            return redirect(url_for("generate"))

        extract_path = None
        if zip_file and zip_file.filename:
            try:
                zip_path = save_uploaded_file(zip_file, work_dir)
                extract_path = os.path.join(work_dir, "poc")
                extract_zip(zip_path, extract_path)
            except Exception as exc:
                logger.exception("ZIP extraction failed")
                flash(f"Unable to extract PoC ZIP: {exc}")
                return redirect(url_for("generate"))

        try:
            output = generate_report(
                excel_path,
                report_type,
                output_name=file_name,
                output_dir=REPORT_FOLDER,
                export_pdf_flag=False,
                screenshot_dir=extract_path,
            )
        except (ReportGenerationError, TemplateLoaderError) as exc:
            logger.exception("Report generation failed")
            flash(str(exc))
            return redirect(url_for("generate"))
        except Exception as exc:
            logger.exception("Unexpected error during report generation")
            flash(f"Unexpected error: {exc}")
            return redirect(url_for("generate"))

        flash(f"Report generated successfully: {os.path.basename(output['docx_path'])}")
        return redirect(url_for("reports"))

    return render_template("generate.html", report_types=list(TEMPLATE_METADATA.keys()))


@app.route("/reports")
@login_required
def reports():
    query = request.args.get("q", "").lower()
    files = []
    if os.path.isdir(REPORT_FOLDER):
        for entry in sorted(
            os.listdir(REPORT_FOLDER),
            key=lambda x: os.path.getmtime(os.path.join(REPORT_FOLDER, x)),
            reverse=True,
        ):
            if query and query not in entry.lower():
                continue
            file_path = os.path.join(REPORT_FOLDER, entry)
            modified_time = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%d-%m-%Y %H:%M:%S")
            files.append({"name": entry, "time": modified_time})
    return render_template("reports.html", files=files, query=query)


@app.route("/download/<filename>")
@login_required
def download(filename: str):
    return send_from_directory(REPORT_FOLDER, filename, as_attachment=True)


@app.route("/suggest")
@login_required
def suggest():
    query = request.args.get("q", "").lower()
    suggestions = []
    if os.path.isdir(REPORT_FOLDER):
        suggestions = [name for name in os.listdir(REPORT_FOLDER) if query in name.lower()]
    return jsonify(suggestions)


@app.route("/logout")
def logout():
    logout_user()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5555)
