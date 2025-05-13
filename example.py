"""
Example usage of the word_file_parser module.
"""

from word_file_parser import DocxParser

def main():
    # Example usage
    try:
        # Initialize the parser with a Word document
        parser = DocxParser("path/to/your/document.docx")
        
        # Parse the document into sections
        sections = parser.parse_sections()
        
        # Print each section with its elements
        print(f"Found {len(sections)} sections:\n")
        for section_title, elements in sections.items():
            print(f"Section: {section_title}")
            print("-" * 50)
            
            # Get formatted text content
            text_content = parser.get_section_text(section_title)
            print(text_content)
            print("\n")
        
        # Save sections to separate files
        parser.save_sections_to_files("output_sections")
        print("\nSections have been saved to the 'output_sections' directory")
        print("Each section's text is saved as a .txt file")
        print("Images are saved as .png files with the naming pattern: section_name_image_N.png")
        
        # Save sections to JSON (with images saved to 'output_images')
        parser.to_json("output_sections/sections.json", image_dir="output_sections/images")
        print("\nSections and elements have been exported to 'output_sections/sections.json'")
        print("Images are saved in 'output_sections/images'")
        
    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main() 