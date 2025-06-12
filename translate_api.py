from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import openai

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

openai.api_key = os.getenv("OPENAI_API_KEY")

def match_terms_to_paragraph(paragraph, terms_df):
    paragraph_lower = paragraph.lower()
    matched = []
    for _, row in terms_df.iterrows():
        term = row['Term'].strip().lower()
        if term in paragraph_lower:
            matched.append(f"{row['Term']} = {row['Translation']}")
    return matched

@app.route('/translate', methods=['POST'])
def translate_file():
    service = request.form.get("service", "Translation_in_Excel")
    uploaded_file = request.files.get('file')
    glossary_file = request.files.get('glossary')

    if not uploaded_file:
        return jsonify({"error": "No file uploaded"}), 400

    filename_raw = uploaded_file.filename
    filename_clean = os.path.splitext(secure_filename(filename_raw))[0]
    file_ext = os.path.splitext(filename_raw)[1].lower()
    input_path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(filename_raw))
    uploaded_file.save(input_path)

    if service == "Translation_in_Excel":
        if file_ext == ".csv":
            paragraphs_df = pd.read_csv(input_path)
        elif file_ext == ".docx":
            doc = Document(input_path)
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            paragraphs_df = pd.DataFrame(paragraphs, columns=["Extracted Paragraphs"])
        else:
            return jsonify({"error": "Unsupported file type for translation"}), 400

        if glossary_file:
            glossary_name = secure_filename(glossary_file.filename)
            glossary_path = os.path.join(app.config["UPLOAD_FOLDER"], glossary_name)
            glossary_file.save(glossary_path)
            doc = Document(glossary_path)
            table = doc.tables[0]
            terms = [row.cells[0].text.strip() for row in table.rows[1:]]
            translations_glossary = [row.cells[1].text.strip() for row in table.rows[1:]]
            glossary_df = pd.DataFrame({"Term": terms, "Translation": translations_glossary})
        else:
            glossary_df = pd.DataFrame(columns=["Term", "Translation"])

        intro_paragraphs = paragraphs_df.iloc[:2, 0].tolist()
        joined_intro = "
".join(intro_paragraphs)
        context_prompt = f"""You are given the beginning of a technical or regulatory document.
Your task is to generate a single clear English sentence that describes the main topic or context of the document.

Content:
{joined_intro}

Context hint:"""

        try:
            context_response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are an expert technical summarizer."},
                    {"role": "user", "content": context_prompt}
                ],
                temperature=0.3,
                max_tokens=50
            )
            context_hint = context_response.choices[0].message.content.strip()
        except Exception:
            context_hint = "[Context generation failed]"

        translations = []
        for paragraph in paragraphs_df.iloc[:, 0]:
            if '----media/' in paragraph:
                translations.append(paragraph.strip())
                continue

            matched_terms = match_terms_to_paragraph(paragraph, glossary_df)
            glossary_text = "
".join(matched_terms) if matched_terms else "[No relevant glossary terms]"
            prompt = f"""[STRICT HAMADA TRANSLATION PROMPT]

Document Context: {context_hint}

Translate the following text from English to Arabic using precise and literal translation.
Do not paraphrase or summarize. Use the glossary below exactly as provided if matching terms are found.
Maintain the original sentence structure and order. Avoid any interpretation or stylistic changes.

Glossary:
{glossary_text}

Text:
{paragraph}"""

            try:
                response = openai.ChatCompletion.create(
                    model="gpt-4",
                    messages=[
                        {"role": "system", "content": "You are a professional English-to-Arabic legal and technical translator."},
                        {"role": "user", "content": prompt.strip()}
                    ],
                    temperature=0.0,
                    max_tokens=1000
                )
                translations.append(response.choices[0].message.content.strip())
            except Exception as e:
                translations.append(f"[Error] {str(e)}")

        paragraphs_df["Translation"] = translations
        excel_output = os.path.join(app.config["UPLOAD_FOLDER"], f"{filename_clean}_translated.xlsx")
        paragraphs_df.to_excel(excel_output, index=False)
        return send_file(excel_output, as_attachment=True)

    elif service == "Convert_to_Word":
        if file_ext != ".xlsx":
            return jsonify({"error": "Expected an Excel file for Word generation"}), 400

        df = pd.read_excel(input_path)
        if "Translation" not in df.columns:
            return jsonify({"error": "Missing 'Translation' column in Excel file"}), 400

        word_output = os.path.join(app.config["UPLOAD_FOLDER"], f"{filename_clean}_translated.docx")
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Simplified Arabic'
        style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Simplified Arabic')
        style.font.size = Pt(14)
        doc.styles['Normal'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

        for text in df["Translation"].fillna(""):
            para = doc.add_paragraph(str(text).strip())
            para.paragraph_format.space_after = Pt(0)

        doc.save(word_output)
        return send_file(word_output, as_attachment=True)

    else:
        return jsonify({"error": "Unknown service type"}), 400

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000")
