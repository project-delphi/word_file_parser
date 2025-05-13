# Word File Parser

A Python library for parsing Word documents (.docx) and splitting them into sections using the `unstructured` library.

## Features

- Parse Word documents into sections based on headings
- Extract section content and metadata
- Save sections to individual text files
- Support for nested sections with heading levels
- Robust handling of document structure

## Installation

1. Clone the repository:

```bash
git clone git@github.com:project-delphi/word_file_parser.git
cd word_file_parser
```

2. Create and activate a virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

```python
from docx_parser import DocxParser

# Initialize parser with a Word document
parser = DocxParser("path/to/document.docx")

# Parse sections
sections = parser.parse_sections()

# Get a specific section
introduction = parser.get_section("Introduction")

# Save all sections to files
parser.save_sections_to_files("output_directory")

# Save a specific section
parser.save_section("Methods", "output_directory")
```

## Development

### Running Tests

```bash
python -m pytest tests/
```

### Code Style

This project uses:

- `black` for code formatting
- `isort` for import sorting
- `flake8` for linting

## License

MIT License
