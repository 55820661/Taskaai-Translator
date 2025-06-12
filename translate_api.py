
import os
import pandas as pd
from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
import tempfile

app = Flask(__name__)
CORS(app)

@app.route("/translate", methods=["POST"])
def translate_file():
    uploaded_file = request.files.get("file")
    if not uploaded_file:
        return {"error": "No file uploaded"}, 400

    output_type = request.form.get("output_type", "docx").lower()
    filename = os.path.splitext(uploaded_file.filename)[0]

    paragraphs = []
    ext = os.path.splitext(uploaded_file.filename)[1].lower()

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, uploaded_file.filename)
        uploaded_file.save(input_path)

        if ext == ".docx":
            doc = Document(input_path)
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        elif ext in [".csv", ".xlsx"]:
            df = pd.read_excel(input_path) if ext == ".xlsx" else pd.read_csv(input_path)
            paragraphs = df.iloc[:, 0].dropna().astype(str).tolist()
        else:
            return {"error": "Unsupported file type"}, 400

        # ترجمة وهمية
        translations = [f"ترجمة: {p}" for p in paragraphs]

        if output_type == "docx":
            word_path = os.path.join(tmpdir, f"{filename}_translated.docx")
            new_doc = Document()
            for t in translations:
                new_doc.add_paragraph(t)
            new_doc.save(word_path)
            return send_file(word_path, as_attachment=True, download_name=f"{filename}_translated.docx")

        elif output_type == "xlsx":
            excel_path = os.path.join(tmpdir, f"{filename}_translated.xlsx")
            df = pd.DataFrame({"Original": paragraphs, "Translated": translations})
            df.to_excel(excel_path, index=False)
            return send_file(excel_path, as_attachment=True, download_name=f"{filename}_translated.xlsx")

        else:
            return {"error": "Invalid output_type"}, 400

if __name__ == "__main__":
    app.run(debug=True)
