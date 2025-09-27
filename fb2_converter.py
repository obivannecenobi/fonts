"""Utilities for converting DOCX documents into FB2 files."""

import os
from pathlib import Path

from docx import Document
from lxml import etree


FB2_NAMESPACE = "http://www.gribuser.ru/xml/fictionbook/2.0"
NSMAP = {None: FB2_NAMESPACE}


def _unique_output_path(destination: Path, original_name: str) -> Path:
    base_name = Path(original_name).stem
    candidate = destination / f"{base_name}.fb2"
    counter = 2
    while candidate.exists():
        candidate = destination / f"{base_name} ({counter}).fb2"
        counter += 1
    return candidate


def convert_docx_to_fb2(docx_path: str, destination_dir: str) -> Path:
    """Convert a DOCX file to FB2 and return the path to the generated file."""
    destination = Path(destination_dir)
    destination.mkdir(parents=True, exist_ok=True)

    document = Document(docx_path)

    root = etree.Element("FictionBook", nsmap=NSMAP)
    body = etree.SubElement(root, "body")
    section = etree.SubElement(body, "section")

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        p_elem = etree.SubElement(section, "p")
        p_elem.text = text

    tree = etree.ElementTree(root)
    output_path = _unique_output_path(destination, os.path.basename(docx_path))
    tree.write(
        str(output_path), encoding="utf-8", xml_declaration=True, pretty_print=True
    )
    return output_path


__all__ = ["convert_docx_to_fb2"]
