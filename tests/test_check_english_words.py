import os
import sys
from pathlib import Path

from docx import Document

sys.path.append(str(Path(__file__).resolve().parents[1]))
from cod import check_english_words


def test_check_english_words(tmp_path):
    path = tmp_path / "english.docx"
    doc = Document()
    doc.add_paragraph("Привет Hello мир")
    doc.add_paragraph("Это Test документ")
    doc.add_paragraph("Другие слова: World, Hello")
    doc.save(path)

    words = check_english_words(str(path))
    assert words == ["Hello", "Test", "World"]

    save_path = tmp_path / "words.txt"
    with open(save_path, "w", encoding="utf-8") as f:
        f.write("\n".join(words))

    assert save_path.read_text(encoding="utf-8") == "Hello\nTest\nWorld"

