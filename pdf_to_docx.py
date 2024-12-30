import os
import re
from pathlib import Path

import mammoth
import markdownify
from docx import Document
from docx.oxml.ns import qn
from docx.parts.image import ImagePart
from pdf2docx import Converter


class WordDocumentParser:
    def __init__(self, file_path: str):
        """
        Initialize the WordDocumentParser with the given file path.

        Parameters:
        file_path (str): The path to the .docx file to be processed.
        """
        self.file_path = Path(file_path)
        self.base_folder = self.file_path.parent / self.file_path.stem
        self.image_folder = self.base_folder / "images"
        self.image_counter = 0
        self.image_map = {}
        self._initialize_document()

    def _initialize_document(self):
        """
        Load the Word document and create the image folder if it doesn't exist.

        Raises:
        Exception: If the document fails to load.
        """
        try:
            self.document = Document(str(self.file_path))
            self.image_folder.mkdir(parents=True, exist_ok=True)
        except Exception as ex:
            raise Exception(f"Failed to load document: {str(ex)}")

    def get_image_filename(self, rel) -> str:
        """
        Generate a filename for an image based on the document name and an incrementing counter.

        Parameters:
        rel: The relationship object for the image.

        Returns:
        str: The generated image filename.
        """
        self.image_counter += 1
        ext = rel.target_part.content_type.split("/")[-1]
        return f"{self.file_path.stem}_{self.image_counter}.{ext}"

    def extract_images(self):
        """
        Extract images from the Word document and save them to the image folder.
        """
        image_rels = {}

        # Collect all internal image relationships
        for rel in self.document.part.rels.values():
            if not rel.is_external and isinstance(rel.target_part, ImagePart):
                image_rels[rel.rId] = rel

        image_order = []

        # Collect the order of image references in the document paragraphs
        for paragraph in self.document.paragraphs:
            for run in paragraph.runs:
                blips = run._element.xpath(".//a:blip")
                if blips:
                    rId = blips[0].get(qn("r:embed"))
                    if rId in image_rels:
                        image_order.append(rId)

        # Extract and save images based on the collected order
        for rId in image_order:
            rel = image_rels[rId]
            image_filename = self.get_image_filename(rel)
            image_path = self.image_folder / image_filename
            self.image_map[rId] = f"images/{image_filename}"

            with open(image_path, "wb") as f:
                f.write(rel.target_part.blob)

    def parse_with_mammoth(self) -> str:
        """
        Convert the Word document to HTML using Mammoth, then convert to Markdown.

        Returns:
        str: The converted Markdown content.

        Raises:
        Exception: If the document fails to convert to HTML.
        """
        try:
            with open(self.file_path, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html = result.value
        except Exception as ex:
            raise Exception(f"Failed to convert document to HTML: {str(ex)}")

        markdown = markdownify.markdownify(html, heading_style="ATX")

        for rId, image_filename in self.image_map.items():
            image_tag = f"![]({image_filename})"
            markdown = re.sub(rf"!\[.*\]\(data:image[^\)]*\)", image_tag, markdown, 1)

        return markdown

    def extract_links(self) -> list:
        """
        Extract external links from the Word document.

        Returns:
        list: A list of tuples containing the relationship ID and the target URL.
        """
        return [
            (rel.rId, rel.target_ref)
            for rel in self.document.part.rels.values()
            if rel.is_external
        ]

    def parse_with_mammoth_and_links(self) -> str:
        """
        Extract images and links, then convert the Word document to Markdown with images and links embedded.

        Returns:
        str: The converted Markdown content with images and links.
        """
        self.extract_images()
        markdown_content = self.parse_with_mammoth()

        for rId, url in self.extract_links():
            # Replace placeholder with the actual link
            markdown_content = re.sub(
                rf'\[.*?\]\(rel="{rId}"\)', f"[{url}]({url})", markdown_content
            )

        return markdown_content

    def save_as_markdown_with_mammoth(self, output_file_path: Path):
        """
        Save the converted Markdown content to a file.

        Parameters:
        output_file_path (Path): The path where the Markdown file will be saved.

        Raises:
        Exception: If the file fails to save.
        """
        markdown_content = self.parse_with_mammoth_and_links()

        try:
            with open(output_file_path, "w", encoding="utf-8") as file:
                file.write(markdown_content)
            print(f"Markdown file saved to {output_file_path}")
        except Exception as ex:
            raise Exception(f"Failed to save Markdown file: {str(ex)}")


def pdf_to_docx(pdf_file):
    """
    Convert a PDF file to DOCX format and save it in the same location as the PDF file.

    Args:
        pdf_file (str): The path to the PDF file to be converted.
    """
    base_name = os.path.splitext(os.path.basename(pdf_file))[0]
    directory = os.path.dirname(pdf_file)

    cv = Converter(pdf_file)

    docx_file = os.path.join(directory, f"{base_name}.docx")
    cv.convert(docx_file)

    cv.close()
    print(f"Conversion completed: '{pdf_file}' to '{docx_file}'")
    return docx_file


def process_file(file_path: str):
    """
    Process the given file, converting PDF to DOCX if necessary, and then parsing it.

    Parameters:
    file_path (str): The path to the file to be processed.
    """
    file_path = Path(file_path)  # Convert the input to a Path object

    if file_path.suffix.lower() == '.pdf':
        file_path = Path(pdf_to_docx(file_path))

    parser = WordDocumentParser(file_path)

    base_folder = file_path.parent / file_path.stem
    output_path = base_folder / f"{file_path.stem}.md"

    try:
        parser.save_as_markdown_with_mammoth(output_path)
    except Exception as e:
        print(f"Error parsing document: {str(e)}")


if __name__ == "__main__":
    file_path = "TestFiles/pdftoword3.pdf"  # Specify your file path here
    process_file(file_path)
