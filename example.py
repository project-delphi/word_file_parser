"""
Example usage of the docx_parser module.
"""

from word_file_parser import DocxParser

def main():
    # Example usage
    try:
        # Initialize the parser with a Word document
        parser = DocxParser("path/to/your/document.docx")
        
        # Parse the document into sections
        sections = parser.parse_sections()
        
        # Print each section
        print(f"Found {len(sections)} sections:\n")
        for section in sections:
            print(f"Section: {section.title}")
            print(f"Level: {section.level}")
            print(f"Content length: {len(section.content)} paragraphs")
            print("-" * 50)
        
        # Save sections to separate files
        parser.save_sections_to_files("output_sections")
        print("\nSections have been saved to the 'output_sections' directory")
        
    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main() 