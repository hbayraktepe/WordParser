import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
import logging

import mammoth
import markdownify
from docx import Document
from docx.oxml.ns import qn
from docx.parts.image import ImagePart
from docx.document import Document as _Document

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


@dataclass
class ImageInfo:
    """Data class for storing image information"""
    filename: str
    content: bytes
    content_type: str


class WordDocumentParser:
    """
    A high-performance Word Document Parser that converts .docx files to Markdown
    with metadata annotations and image extraction capabilities.
    """

    # Class level constants
    HEADING_PATTERN = re.compile(r'^#+\s')
    LIST_BULLET_PATTERN = re.compile(r'^\s*[\*\-\+]\s')
    LIST_NUMBER_PATTERN = re.compile(r'^\s*\d+\.\s')
    TABLE_SEPARATOR_PATTERN = re.compile(r'\|\s*[-:]+\s*\|')
    IMAGE_PATTERN = re.compile(r'!\[.*?\]\(data:image[^\)]*\)')

    def __init__(self, docx_file_path: str, extract_images: bool = True):
        """
        Initialize the WordDocumentParser with configuration and setup paths.
        """
        self.file_path = Path(docx_file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"Document not found: {docx_file_path}")

        self.extract_images_option = extract_images
        self.base_folder = self.file_path.parent / self.file_path.stem
        self.image_folder = self.base_folder / "images"
        self.image_counter = 0
        self.image_map: Dict[str, ImageInfo] = {}

        self._initialize()

    def _initialize(self) -> None:
        """Set up the parser's initial state and resources."""
        try:
            self.document = Document(str(self.file_path))
            if self.extract_images_option:
                self.image_folder.mkdir(parents=True, exist_ok=True)
            logger.info(f"Initialized parser for: {self.file_path}")
        except Exception as ex:
            logger.error(f"Initialization failed: {ex}")
            raise

    def _process_image(self, element: Any, rId: str) -> Optional[ImageInfo]:
        """Process a single image element from the document."""
        try:
            rel = self.document.part.rels[rId]
            if not isinstance(rel.target_part, ImagePart):
                return None

            extension = rel.target_part.content_type.split('/')[-1]
            filename = f"{self.file_path.stem}_{self.image_counter}.{extension}"

            return ImageInfo(
                filename=filename,
                content=rel.target_part.blob,
                content_type=rel.target_part.content_type
            )
        except Exception as ex:
            logger.warning(f"Failed to process image {rId}: {ex}")
            return None

    def _collect_images(self) -> None:
        """Collect and store all images from the document."""
        if not self.extract_images_option:
            return

        image_rels = {
            rel.rId: rel for rel in self.document.part.rels.values()
            if not rel.is_external and isinstance(rel.target_part, ImagePart)
        }

        for paragraph in self.document.paragraphs:
            for run in paragraph.runs:
                blips = run._element.xpath(".//a:blip")
                if not blips:
                    continue

                rId = blips[0].get(qn("r:embed"))
                if rId not in image_rels:
                    continue

                image_info = self._process_image(run, rId)
                if image_info:
                    self.image_map[rId] = image_info
                    self.image_counter += 1

    def _save_images(self) -> None:
        """Save collected images to the filesystem."""
        if not self.extract_images_option:
            return

        for rId, image_info in self.image_map.items():
            image_path = self.image_folder / image_info.filename
            try:
                image_path.write_bytes(image_info.content)
                logger.debug(f"Saved image: {image_path}")
            except Exception as ex:
                logger.error(f"Failed to save image {image_path}: {ex}")

    @staticmethod
    def _is_list_item(line: str) -> bool:
        """Check if the line is a list item."""
        stripped = line.lstrip()
        return bool(stripped and (
                stripped.startswith(('*', '-', '+')) or
                re.match(r'^\d+\.', stripped)
        ))

    @staticmethod
    def _get_indent_level(line: str) -> int:
        """Get the indentation level of a line."""
        return len(line) - len(line.lstrip())

    def _process_list_block(self, lines: List[str]) -> List[str]:
        """Process a block of list items."""
        if not lines:
            return []

        # Map indentation levels to normalized levels
        indent_to_level = {}
        current_level = 0

        # First pass: collect indentation levels
        for line in lines:
            if self._is_list_item(line):
                indent = self._get_indent_level(line)
                if indent not in indent_to_level:
                    indent_to_level[indent] = current_level
                    current_level += 1

        # Sort indents for consistent level mapping
        indent_levels = sorted(indent_to_level.keys())

        # Second pass: normalize indentation and markers
        result = []
        for line in lines:
            stripped = line.lstrip()
            if not stripped:
                result.append(line)
                continue

            indent = self._get_indent_level(line)
            # Find the closest indent level
            closest_indent = min(indent_levels, key=lambda x: abs(x - indent))
            level = indent_to_level[closest_indent]

            # Process list items
            if self._is_list_item(line):
                content = stripped.lstrip("*-+").lstrip()
                if re.match(r'^\d+\.', stripped):
                    content = re.sub(r'^\d+\.\s*', '', stripped)
                result.append("  " * level + "* " + content)
            else:
                result.append("  " * (level + 1) + stripped)

        # Add a single metadata comment at the end of the list block
        result.append("<!-- Type: List Block -->")

        return result

    def _process_nested_lists(self, lines: List[str]) -> List[str]:
        """Process and properly align nested list items."""
        result = []
        list_block = []
        in_list = False

        for line in lines:
            stripped = line.lstrip()

            if self._is_list_item(line):
                if not in_list:
                    if list_block:
                        result.extend(self._process_list_block(list_block))
                        list_block = []
                    in_list = True
                list_block.append(line)
            else:
                if in_list:
                    if stripped:
                        list_block.append(line)
                    else:
                        if list_block:
                            result.extend(self._process_list_block(list_block))
                            list_block = []
                        in_list = False
                        result.append(line)
                else:
                    processed_content, _ = self._add_metadata(line, in_list=False)
                    result.extend(processed_content)

        if list_block:
            result.extend(self._process_list_block(list_block))

        return result

    def _add_metadata(self, line: str, in_list: bool = False) -> Tuple[List[str], Optional[str]]:
        """Add appropriate metadata comment for a line of markdown."""
        line = line.rstrip()
        if not line:
            return [line], None

        result = []
        metadata = None

        if self.HEADING_PATTERN.match(line):
            level = len(line.split()[0])
            result.extend([line, f"<!-- Type: Heading {level} -->"])
            metadata = f"heading_{level}"
        elif self._is_list_item(line):
            result.append(line)
            metadata = "list_item"
        elif line.strip().startswith('|') and '-|-' in line:
            result.extend([line, "<!-- Type: Table -->"])
            metadata = "table_separator"
        elif line.strip().startswith('|'):
            result.append(line)
            metadata = "table_row"
        elif '![' in line and '](' in line:
            result.extend([line, "<!-- Type: Image -->"])
            metadata = "image"
        elif line.strip() and not in_list:
            result.extend([line, "<!-- Type: Text Body -->"])
            metadata = "text"
        else:
            result.append(line)

        return result, metadata

    def convert_to_markdown(self) -> str:
        """Convert the Word document to markdown with metadata."""
        try:
            self._collect_images()
            self._save_images()

            with open(self.file_path, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                raw_markdown = markdownify.markdownify(result.value, heading_style="ATX")

            # Process lists first
            lines = raw_markdown.splitlines()
            processed_lines = self._process_nested_lists(lines)

            # Then process metadata and tables
            markdown_lines = []
            current_table = []
            in_table = False

            for line in processed_lines:
                if '<!-- Type:' in line:  # Already processed metadata
                    markdown_lines.append(line)
                    continue

                if line.strip().startswith('|'):
                    if not in_table:
                        current_table = [line]
                        in_table = True
                    else:
                        current_table.append(line)
                else:
                    if in_table:
                        markdown_lines.extend(current_table)
                        current_table = []
                        in_table = False
                        markdown_lines.append("")

                    markdown_lines.append(line)

            if current_table:
                markdown_lines.extend(current_table)
                markdown_lines.append("")

            content = '\n'.join(markdown_lines)

            # Process images and links
            for rId, image_info in self.image_map.items():
                image_tag = f"![Image](images/{image_info.filename})"
                content = self.IMAGE_PATTERN.sub(image_tag, content, 1)

            for rel in self.document.part.rels.values():
                if rel.is_external:
                    content = re.sub(
                        rf'\[([^\]]*)\]\(rel="{rel.rId}"\)',
                        f"[\\1]({rel.target_ref})",
                        content
                    )

            return content

        except Exception as ex:
            logger.error(f"Conversion failed: {ex}")
            raise

    def save_markdown(self, output_path: Optional[Path] = None) -> None:
        """Save the converted markdown content to a file."""
        if output_path is None:
            output_path = self.base_folder / f"{self.file_path.stem}.md"

        try:
            content = self.convert_to_markdown()
            output_path.write_text(content, encoding='utf-8')
            logger.info(f"Successfully saved markdown to: {output_path}")
        except Exception as ex:
            logger.error(f"Failed to save markdown: {ex}")
            raise


if __name__ == "__main__":
    file_path = "TestFiles/complex_test1.docx"
    try:
        parser = WordDocumentParser(file_path)
        parser.save_markdown()
        logger.info("Document conversion completed successfully")
    except Exception as e:
        logger.error(f"Error processing document: {str(e)}")