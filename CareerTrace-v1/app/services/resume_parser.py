from pathlib import Path
from docx import Document
from pypdf import PdfReader

def extract_resume(path: Path) -> str:
    if not path.exists():
        return ""
    if path.suffix.lower() == ".docx":
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    if path.suffix.lower() == ".pdf":
        reader = PdfReader(str(path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    raise ValueError("Supported resume formats: .docx and .pdf")
