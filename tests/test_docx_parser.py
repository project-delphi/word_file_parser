import os
import pytest
from pathlib import Path
from word_file_parser import DocxParser

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
    assert isinstance(docx_parser.sections, dict) or isinstance(docx_parser.sections, list)
    assert len(docx_parser.sections) == 0

def test_parse_sections(docx_parser):
    """Test parsing sections from a docx file."""
    docx_parser.parse_sections()
    assert len(docx_parser.sections) == 5  # We have 5 sections in our test file
    section_titles = [section.title for section in docx_parser.sections]
    for expected in ['Introduction', 'Methods', 'Results', 'Discussion', 'Conclusion']:
        assert expected in section_titles

def test_get_section(docx_parser):
    """Test retrieving a section by title."""
    # Test with non-existent section
    assert docx_parser.get_section("NonExistentSection") is None
    
    # Parse sections first
    docx_parser.parse_sections()
    
    # Test with existing sections
    for section in ['Introduction', 'Methods', 'Results', 'Discussion', 'Conclusion']:
        content = docx_parser.get_section(section)
        assert content is not None
        assert isinstance(content, str)
        assert len(content) > 0

def test_save_section(tmp_path, docx_parser):
    """Test saving a section to a file."""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    
    # Parse sections first
    docx_parser.parse_sections()
    
    # Test saving non-existent section
    with pytest.raises(ValueError):
        docx_parser.save_section("NonExistentSection", str(output_dir))
    
    # Test saving existing sections
    for section in ['Introduction', 'Methods', 'Results', 'Discussion', 'Conclusion']:
        docx_parser.save_section(section, str(output_dir))
        output_file = output_dir / f"{section}.txt"
        assert output_file.exists()
        content = output_file.read_text()
        # Compare to the string representation of the section
        section_obj = docx_parser.get_section_by_title(section)
        assert content == str(section_obj)

def test_parse_docx(sample_docx_path, tmp_path):
    """Test the convenience function for parsing docx files."""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    
    # Test with non-existent file
    with pytest.raises(FileNotFoundError):
        DocxParser.parse_docx("non_existent.docx", str(output_dir))
    
    # Test with real file
    DocxParser.parse_docx(sample_docx_path, str(output_dir))
    assert len(list(output_dir.glob("*.txt"))) == 5  # We should have 5 text files
    for section in ['Introduction', 'Methods', 'Results', 'Discussion', 'Conclusion']:
        output_file = output_dir / f"{section}.txt"
        assert output_file.exists()
        content = output_file.read_text()
        assert len(content) > 0 