import sys
from pathlib import Path

from docx import Document

sys.path.append(str(Path(__file__).resolve().parents[1]))
from cod import find_missing_chapters  # noqa: E402


def _create_document(path, headings):
    document = Document()
    for heading in headings:
        document.add_paragraph(f"Глава {heading}")
        document.add_paragraph("Content")
    document.save(path)


def test_missing_simple_numbering(tmp_path):
    doc_path = tmp_path / "simple.docx"
    _create_document(doc_path, ["1", "2", "4"])

    missing = find_missing_chapters(str(doc_path))

    assert missing == ["Глава 3"]


def test_missing_decimal_numbering(tmp_path):
    doc_path = tmp_path / "decimal.docx"
    _create_document(doc_path, ["1.1", "1.2", "2.1", "3.1"])

    missing = find_missing_chapters(str(doc_path))

    assert missing == ["Глава 2.2"]


def test_no_missing_chapters(tmp_path):
    doc_path = tmp_path / "complete.docx"
    _create_document(doc_path, ["1", "2", "3"])

    missing = find_missing_chapters(str(doc_path))

    assert missing == []

