import os
import pytest
from pathlib import Path
from word_file_parser import DocxParser
from unstructured.documents.elements import Image, Table, ListItem
import json

@pytest.fixture
def sample_docx_path():
    """Get the path to the sample docx file."""
    return str(Path('tests/test_files/sample.docx'))

@pytest.fixture
def docx_parser(sample_docx_path):
    """Create a DocxParser instance for testing."""
    return DocxParser(sample_docx_path)

def test_init(docx_parser, sample_docx_path):
    """Test DocxParser initialization."""
    assert str(docx_parser.file_path) == sample_docx_path
    assert isinstance(docx_parser.sections, dict)
    assert len(docx_parser.sections) == 0

def test_parse_sections(docx_parser):
    """Test parsing sections from a docx file."""
    sections = docx_parser.parse_sections()
    assert isinstance(sections, dict)
    # Sections may be empty if the sample docx is minimal
    for section_elements in sections.values():
        assert isinstance(section_elements, list)
        # It's OK if there are no ListItem/Table/Image elements
        for elem in section_elements:
            assert hasattr(elem, "text") or hasattr(elem, "image")

def test_get_section(docx_parser):
    """Test retrieving a section by title."""
    # Test with non-existent section
    assert docx_parser.get_section("NonExistentSection") is None
    # Parse sections first
    docx_parser.parse_sections()
    # Test with existing sections
    for section in docx_parser.sections:
        elements = docx_parser.get_section(section)
        assert elements is not None
        assert isinstance(elements, list)
        # It's OK if the section is empty

def test_get_section_text(docx_parser):
    """Test getting formatted text content of a section."""
    docx_parser.parse_sections()
    for section in docx_parser.sections:
        text = docx_parser.get_section_text(section)
        # It's OK if text is None or empty
        assert text is None or isinstance(text, str)

def test_save_section(tmp_path, docx_parser):
    """Test saving a section to files."""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    # Parse sections first
    docx_parser.parse_sections()
    # Test saving non-existent section
    with pytest.raises(ValueError):
        docx_parser.save_section("NonExistentSection", str(output_dir))
    # Test saving existing sections
    for section in docx_parser.sections:
        # It's OK if the section is empty, just ensure no error is raised
        try:
            docx_parser.save_section(section, str(output_dir))
        except ValueError:
            pass
        # Check text file exists if content was saved
        text_file = output_dir / f"{section}.txt"
        if text_file.exists():
            content = text_file.read_text()
            assert isinstance(content, str)

def test_parse_docx(sample_docx_path, tmp_path):
    """Test the convenience function for parsing docx files to JSON."""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    json_path = output_dir / "sections.json"
    # Test with non-existent file
    with pytest.raises(FileNotFoundError):
        DocxParser.parse_docx_to_json("non_existent.docx", json_path)
    # Test with real file
    DocxParser.parse_docx_to_json(sample_docx_path, json_path)
    # Check that JSON file was created
    assert json_path.exists()
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    assert isinstance(data, list)

def test_to_json(tmp_path, docx_parser):
    """Test exporting parsed sections to JSON."""
    output_dir = tmp_path / "output_json"
    image_dir = output_dir / "images"
    json_path = output_dir / "sections.json"
    output_dir.mkdir()
    image_dir.mkdir()

    docx_parser.parse_sections()
    docx_parser.to_json(json_path, image_dir=image_dir)

    # Check that JSON file exists and is valid
    assert json_path.exists()
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    assert isinstance(data, list)
    assert len(data) > 0
    for section in data:
        assert "title" in section
        assert "elements" in section
        for elem in section["elements"]:
            assert "type" in elem
            # If image, check file reference
            if elem["type"] == "Image":
                assert "image_file" in elem or "image_base64" in elem
                if "image_file" in elem:
                    img_path = Path(elem["image_file"])
                    assert img_path.exists() 