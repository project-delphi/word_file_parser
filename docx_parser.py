"""
Module for parsing and splitting Word documents into sections using the unstructured library.
"""

from typing import List, Optional
from dataclasses import dataclass
from pathlib import Path

from unstructured.partition.auto import partition
from unstructured.documents.elements import Title, NarrativeText, ListItem


@dataclass
class Section:
    """Represents a section in a document with its title and content."""
    title: str
    content: List[str]
    level: int = 1

    def __str__(self) -> str:
        """String representation of the section."""
        return f"{'#' * self.level} {self.title}\n\n" + "\n".join(self.content)


class DocxParser:
    """Parser for Word documents that splits them into sections using unstructured."""

    def __init__(self, file_path: str):
        """
        Initialize the parser with a Word document.

        Args:
            file_path (str): Path to the Word document
        """
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"Document not found: {file_path}")
        
        # Use the partition function which automatically detects file type
        self.elements = partition(filename=str(self.file_path))
        self.sections: List[Section] = []

    def _get_heading_level(self, element: Title) -> int:
        """
        Determine the heading level from a Title element.
        """
        style_name = getattr(element.metadata, "style_name", "")
        if style_name.lower().startswith("heading"):
            try:
                return int(style_name[-1])
            except (ValueError, KeyError):
                pass
        return 1

    def parse_sections(self) -> List[Section]:
        """
        Parse the document and split it into sections based on headings.

        Returns:
            List[Section]: List of sections with their titles and content
        """
        current_section = None
        current_content = []
        self.sections = []
        current_level = 1
        title_skipped = False

        for element in self.elements:
            if isinstance(element, Title):
                if not title_skipped:
                    # Always skip the first Title element (document title)
                    title_skipped = True
                    continue
                # Save previous section if it exists
                if current_section:
                    self.sections.append(Section(
                        title=current_section,
                        content=current_content,
                        level=current_level
                    ))
                # Start new section
                current_section = str(element)
                current_content = []
                current_level = self._get_heading_level(element)
            elif isinstance(element, (NarrativeText, ListItem)):
                if current_section is None:
                    # Handle content before first heading
                    current_section = "Introduction"
                current_content.append(str(element))
        # Add the last section
        if current_section:
            self.sections.append(Section(
                title=current_section,
                content=current_content,
                level=current_level
            ))
        return self.sections

    def get_section_by_title(self, title: str) -> Optional[Section]:
        """
        Get a specific section by its title.

        Args:
            title (str): The title of the section to find

        Returns:
            Optional[Section]: The found section or None if not found
        """
        for section in self.sections:
            if section.title.lower() == title.lower():
                return section
        return None

    def save_sections_to_files(self, output_dir: str) -> None:
        """
        Save each section to a separate text file.

        Args:
            output_dir (str): Directory to save the section files
        """
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        for section in self.sections:
            # Create a safe filename from the section title
            safe_title = "".join(c if c.isalnum() else "_" for c in section.title)
            file_path = output_path / f"{safe_title}.txt"
            
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(str(section))

    def get_section(self, title: str) -> Optional[str]:
        """
        Get the content of a section by its title (for test compatibility).
        """
        section = self.get_section_by_title(title)
        if section:
            return "\n".join(section.content)
        return None

    def save_section(self, title: str, output_dir: str) -> None:
        """
        Save a single section by title to a text file.
        """
        section = self.get_section_by_title(title)
        if not section:
            raise ValueError(f"Section '{title}' not found.")
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        safe_title = "".join(c if c.isalnum() else "_" for c in section.title)
        file_path = output_path / f"{safe_title}.txt"
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(str(section))

    @classmethod
    def parse_docx(cls, file_path: str, output_dir: str = None):
        """
        Parse a docx file and optionally save sections to files (for test compatibility).
        """
        parser = cls(file_path)
        parser.parse_sections()
        if output_dir:
            parser.save_sections_to_files(output_dir)
        return parser.sections


def parse_docx(file_path: str) -> List[Section]:
    """
    Convenience function to parse a Word document and return its sections.

    Args:
        file_path (str): Path to the Word document

    Returns:
        List[Section]: List of sections in the document
    """
    parser = DocxParser(file_path)
    return parser.parse_sections() 