"""
Module for parsing and splitting Word documents into sections using the unstructured library.
"""

from typing import List, Optional, Dict, Union
from dataclasses import dataclass
from pathlib import Path
import json
import base64

from unstructured.partition.auto import partition
from unstructured.documents.elements import Title, NarrativeText, ListItem, Table, Image, FigureCaption, Text, Element


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
    """Parser for Word documents that splits them into sections."""
    
    def __init__(self, file_path: Union[str, Path]):
        """Initialize the parser with a Word document path.
        
        Args:
            file_path: Path to the Word document
        """
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        self.sections: Dict[str, List[Element]] = {}
        self._current_section: Optional[str] = None
        self._current_elements: List[Element] = []
    
    def parse_sections(self) -> Dict[str, List[Element]]:
        """Parse the document into sections based on headings.
        
        Returns:
            Dictionary mapping section titles to their content elements
        """
        elements = partition(filename=str(self.file_path))
        
        for element in elements:
            if isinstance(element, Title):
                # Save previous section if it exists
                if self._current_section is not None:
                    self.sections[self._current_section] = self._current_elements
                
                # Start new section
                self._current_section = element.text
                self._current_elements = []
            else:
                if self._current_section is not None:
                    self._current_elements.append(element)
        
        # Save the last section
        if self._current_section is not None:
            self.sections[self._current_section] = self._current_elements
        
        return self.sections
    
    def get_section(self, title: str) -> Optional[List[Element]]:
        """Get a section by its title.
        
        Args:
            title: The section title to retrieve
            
        Returns:
            List of elements in the section, or None if not found
        """
        if not self.sections:
            self.parse_sections()
        return self.sections.get(title)
    
    def get_section_text(self, title: str) -> Optional[str]:
        """Get the text content of a section.
        
        Args:
            title: The section title to retrieve
            
        Returns:
            Formatted text content of the section, or None if not found
        """
        elements = self.get_section(title)
        if not elements:
            return None
            
        text_parts = []
        for element in elements:
            if isinstance(element, (NarrativeText, ListItem, Text)):
                text_parts.append(element.text)
            elif isinstance(element, Table):
                text_parts.append("\nTable:\n" + element.text)
            elif isinstance(element, Image):
                text_parts.append("\n[Image]")
                if element.caption:
                    text_parts.append(f"Caption: {element.caption}")
            elif isinstance(element, FigureCaption):
                text_parts.append(f"\nCaption: {element.text}")
        
        return "\n".join(text_parts)
    
    def save_section(self, title: str, output_dir: Union[str, Path]) -> None:
        """Save a section to a text file.
        
        Args:
            title: The section title to save
            output_dir: Directory to save the section file
        """
        elements = self.get_section(title)
        if not elements:
            raise ValueError(f"Section not found: {title}")
            
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Save text content
        text_content = self.get_section_text(title)
        if text_content:
            text_file = output_dir / f"{title}.txt"
            text_file.write_text(text_content)
        
        # Save images if any
        for i, element in enumerate(elements):
            if isinstance(element, Image) and element.image:
                image_file = output_dir / f"{title}_image_{i+1}.png"
                with open(image_file, 'wb') as f:
                    f.write(element.image)
    
    def save_sections_to_files(self, output_dir: Union[str, Path]) -> None:
        """Save all sections to text files.
        
        Args:
            output_dir: Directory to save the section files
        """
        if not self.sections:
            self.parse_sections()
            
        for title in self.sections:
            self.save_section(title, output_dir)
    
    def to_json(self, output_path: Union[str, Path], image_dir: Optional[Union[str, Path]] = None) -> None:
        """Export the parsed sections and their elements to a JSON file.

        Args:
            output_path: Path to the output JSON file
            image_dir: Directory to save images (if not already saved). If None, images are not saved.
        """
        if not self.sections:
            self.parse_sections()
        
        output = []
        for section_title, elements in self.sections.items():
            section_data = {
                "title": section_title,
                "elements": []
            }
            for i, element in enumerate(elements):
                elem_type = type(element).__name__
                elem_data = {"type": elem_type}
                if hasattr(element, "text"):
                    elem_data["text"] = element.text
                if hasattr(element, "caption") and element.caption:
                    elem_data["caption"] = element.caption
                if elem_type == "Table":
                    elem_data["table_text"] = element.text
                if elem_type == "Image" and getattr(element, "image", None):
                    if image_dir:
                        image_dir_path = Path(image_dir)
                        image_dir_path.mkdir(parents=True, exist_ok=True)
                        image_file = image_dir_path / f"{section_title}_image_{i+1}.png"
                        with open(image_file, 'wb') as f:
                            f.write(element.image)
                        elem_data["image_file"] = str(image_file)
                    else:
                        # Optionally, base64 encode the image
                        elem_data["image_base64"] = base64.b64encode(element.image).decode("utf-8")
                section_data["elements"].append(elem_data)
            output.append(section_data)
        
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(output, f, indent=2, ensure_ascii=False)

    @classmethod
    def parse_docx_to_json(cls, file_path: Union[str, Path], json_path: Union[str, Path], image_dir: Optional[Union[str, Path]] = None) -> None:
        """Convenience function to parse a docx file and output JSON.

        Args:
            file_path: Path to the Word document
            json_path: Path to the output JSON file
            image_dir: Directory to save images (optional)
        """
        parser = cls(file_path)
        parser.to_json(json_path, image_dir=image_dir)


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