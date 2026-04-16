from __future__ import annotations

import tempfile
from pathlib import Path

from flask import Flask, render_template, request
from werkzeug.utils import secure_filename

from thesis_checker import analyze_docx

ALLOWED_EXTENSIONS = {".docx"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def index():
    report = None
    error = None
    if request.method == "POST":
        uploaded = request.files.get("file")
        if not uploaded or not uploaded.filename:
            error = "\u8acb\u5148\u9078\u64c7\u4e00\u500b .docx \u6a94\u6848\u3002"
        elif not allowed_file(uploaded.filename):
            error = "\u76ee\u524d\u53ea\u652f\u63f4 .docx \u683c\u5f0f\u3002"
        else:
            safe_name = secure_filename(uploaded.filename) or "uploaded.docx"
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir) / safe_name
                uploaded.save(temp_path)
                report = analyze_docx(temp_path)
    return render_template("index.html", report=report, error=error)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
