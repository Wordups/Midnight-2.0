import os
import tempfile
from docx import Document
from groq import Groq
from hps_policy_migration_builder import build_policy_document

def get_api_key():
    return os.getenv("GROQ_API_KEY")

def parse_policy_data(raw_output: str):
    if "POLICY_DATA = {" in raw_output:
        raw_output = raw_output[raw_output.index("POLICY_DATA = {"):]
    namespace = {}
    exec(raw_output, {}, namespace)
    return namespace.get("POLICY_DATA")

def extract_text_from_docx(file_obj):
    doc = Document(file_obj)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

def get_uploaded_text(filename, file_bytes):
    from io import BytesIO
    if filename.endswith(".docx"):
        return extract_text_from_docx(BytesIO(file_bytes))
    return file_bytes.decode("utf-8", errors="ignore")

def save_logo_bytes(filename, file_bytes):
    path = os.path.join(tempfile.gettempdir(), filename)
    with open(path, "wb") as f:
        f.write(file_bytes)
    return path

def run_llm_transform(source_text, template_name):
    client = Groq(api_key=get_api_key())

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": source_text}],
    )

    raw = response.choices[0].message.content
    data = parse_policy_data(raw)
    data["template_name"] = template_name
    return data

def build_output_doc(policy_data, logo_path=None):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        path = tmp.name

    build_policy_document(policy_data, path, logo_path=logo_path)

    with open(path, "rb") as f:
        data = f.read()

    return "output.docx", data
