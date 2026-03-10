"""
Flask API for PDF flattening and compression.
"""
import os
import tempfile
import uuid
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from spool_optimizer import DocumentSpoolOptimizer

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB
app.config["UPLOAD_FOLDER"] = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {"pdf"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/process", methods=["POST"])
def process_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only PDF files are allowed"}), 400

    dpi = request.form.get("dpi", 100, type=int)
    dpi = max(72, min(300, dpi))

    job_id = str(uuid.uuid4())
    input_path = Path(app.config["UPLOAD_FOLDER"]) / f"{job_id}_input.pdf"
    output_path = Path(app.config["UPLOAD_FOLDER"]) / f"{job_id}_output.pdf"

    try:
        file.save(input_path)

        optimizer = DocumentSpoolOptimizer(dpi=dpi)
        success = optimizer.process_document(input_path, output_path)

        input_path.unlink(missing_ok=True)

        if not success:
            output_path.unlink(missing_ok=True)
            return jsonify({"error": "Failed to process PDF. The file may be corrupted, password-protected, or invalid."}), 422

        return jsonify({
            "success": True,
            "download_id": job_id,
            "original_name": secure_filename(file.filename),
        })
    except MemoryError:
        input_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)
        return jsonify({"error": "File is too large to process. Insufficient memory."}), 413
    except Exception as e:
        input_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)
        return jsonify({"error": f"An unexpected error occurred: {str(e)}"}), 500


@app.route("/api/download/<job_id>")
def download_pdf(job_id):
    output_path = Path(app.config["UPLOAD_FOLDER"]) / f"{job_id}_output.pdf"
    if not output_path.exists():
        return jsonify({"error": "File not found or expired"}), 404

    try:
        return send_file(
            output_path,
            as_attachment=True,
            download_name="flattened.pdf",
            mimetype="application/pdf",
        )
    finally:
        output_path.unlink(missing_ok=True)


if __name__ == "__main__":
    import os
    host = os.environ.get("FLASK_HOST", "127.0.0.1")
    port = int(os.environ.get("FLASK_PORT", 5000))
    debug = os.environ.get("FLASK_ENV", "development") != "production"
    app.run(host=host, port=port, debug=debug)
