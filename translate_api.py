
from flask import Flask, request, jsonify
from flask_cors import CORS

import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
import tempfile
import openai

app = Flask(__name__)
CORS(app)

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
    uploaded_file = request.files.get('file')
    glossary_file = request.files.get('glossary')

    if not uploaded_file:
        return jsonify({"error": "No file uploaded"}), 400

    filename = os.path.splitext(uploaded_file.filename)[0]
    file_ext = os.path.splitext(uploaded_file.filename)[1].lower()

    with tempfile.TemporaryDirectory() as tmpdirname:
        input_path = os.path.join(tmpdirname, uploaded_file.filename)
        uploaded_file.save(input_path)

        if file_ext == ".csv":
            paragraphs_df = pd.read_csv(input_path)
        elif file_ext == ".docx":
            doc = Document(input_path)
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            paragraphs_df = pd.DataFrame(paragraphs, columns=["Extracted Paragraphs"])
        else:
            return jsonify({"error": "Unsupported file type"}), 400

        if glossary_file:
            glossary_path = os.path.join(tmpdirname, glossary_file.filename)
            glossary_file.save(glossary_path)
            doc = Document(glossary_path)
            table = doc.tables[0]
            terms = [row.cells[0].text.strip() for row in table.rows[1:]]
            translations_glossary = [row.cells[1].text.strip() for row in table.rows[1:]]
            glossary_df = pd.DataFrame({"Term": terms, "Translation": translations_glossary})
        else:
            glossary_df = pd.DataFrame(columns=["Term", "Translation"])

        intro_paragraphs = paragraphs_df.iloc[:2, 0].tolist()
        joined_intro = "\n".join(intro_paragraphs)
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
        except Exception as e:
            context_hint = "[Context generation failed]"

        translations = []
        for i, paragraph in enumerate(paragraphs_df.iloc[:, 0]):
            if '----media/' in paragraph:
                translations.append(paragraph.strip())
                continue

            matched_terms = match_terms_to_paragraph(paragraph, glossary_df)
            glossary_text = "\n".join(matched_terms) if matched_terms else "[No relevant glossary terms]"
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

        excel_output = os.path.join(tmpdirname, f"{filename}_translated.xlsx")
        word_output = os.path.join(tmpdirname, f"{filename}_translated.docx")

        paragraphs_df.to_excel(excel_output, index=False)

        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Simplified Arabic'
        style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Simplified Arabic')
        style.font.size = Pt(14)
        section = doc.sections[0]
        section.right_margin = section.left_margin
        doc.styles['Normal'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

        for text in translations:
            if isinstance(text, str):
                para = doc.add_paragraph(text.strip())
                para.paragraph_format.space_after = Pt(0)

        doc.save(word_output)

        return {
            "excel_file": f"{filename}_translated.xlsx",
            "word_file": f"{filename}_translated.docx",
            "message": "✅ Translation complete"
        }

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
