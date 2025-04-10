from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
from docx2pdf import convert  # Word to PDF
from pdf2docx import Converter  # PDF to Word

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert_file():
    uploaded_file = request.files["file"]
    conversion_type = request.form["conversion"]

    if not uploaded_file:
        return "No file uploaded."

    filename = secure_filename(uploaded_file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    uploaded_file.save(filepath)

    # Output file path
    name, ext = os.path.splitext(filename)
    output_path = ""

    try:
        if conversion_type == "word_to_pdf" and ext in [".docx", ".doc"]:
            output_path = os.path.join(CONVERTED_FOLDER, f"{name}.pdf")
            convert(filepath, output_path)

        elif conversion_type == "pdf_to_word" and ext == ".pdf":
            output_path = os.path.join(CONVERTED_FOLDER, f"{name}.docx")
            cv = Converter(filepath)
            cv.convert(output_path, start=0, end=None)
            cv.close()

        else:
            return "Unsupported file type or conversion option."

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Conversion failed: {str(e)}"

if __name__ == "__main__":
    app.run(debug=True)
