import re
import easyocr
import logging
import pdfplumber
import numpy as np
from PIL import Image
from pdf2image import convert_from_path


class PDFProcessor:
    def __init__(
        self,
        poppler_path: str | None = None,
        dpi: int = 600,
        lang_list: list = ["ja", "en"],
        logger_name: str = __name__,
    ):
        self.poppler_path = poppler_path
        self.reader = easyocr.Reader(
            lang_list=lang_list,
        )
        self.dpi = dpi
        self.logger = logging.getLogger(logger_name)

        # Note: Only the first encountered keyword will be processed
        self.keywords = [
            "建築地住所",
            "建築地",
            "申請地",
            "現場地図",
            "建設地",
        ]

        # Custom bounding box adjustments (left, top, right, bottom)
        self.bbox_adjustments = {
            "建築地住所": (-50, -20, 3300, 35),
            "建築地": (-50, -20, 3300, 35),
            "申請地": (-50, -20, 3300, 35),
            "現場地図": (-50, -20, 3300, 35),
            "建設地": (-50, -20, 3300, 35),
        }

    def clean_address(self, address):
        """
        Clean and standardize the extracted address.
        Removes numbers at the beginning and keeps only the address until
        encountering either a whitespace or a number.
        """
        # Remove specific keywords from the beginning of the address
        for keyword in self.keywords:
            if address.startswith(keyword):
                address = address.replace(keyword, "").strip()

        # Remove any excess whitespace
        address = re.sub(r"\s+", " ", address)

        # Remove numbers at the beginning of the address
        address = re.sub(r"^[\d\s-]+", "", address)

        # Keep only the address until the first whitespace or number
        match = re.search(r"^([^\d\s]+)", address)
        if match:
            address = match.group(1)

        return address.strip()

    def process_image_ocr(self, image_np) -> tuple[bool, dict[str, str] | None]:
        """
        Process a single image with OCR and save screenshots of first keyword occurrences.

        Args:
            image_np (numpy.ndarray): Input image as a numpy array
        Returns:
            tuple: (success_flag, keyword_addresses)
        """
        # Perform OCR on the entire image
        results = self.reader.readtext(
            image=image_np,
            detail=1,
            paragraph=False,
        )
        keyword_found = False
        first_keyword_pos = None
        first_keyword = None
        # First pass: find the first keyword
        for result in results:
            if len(result) < 2:
                continue
            bbox, text = result[:2]

            # Check for the first keyword occurrence
            for keyword in self.keywords:
                if keyword in text:
                    # If no keyword has been found yet, record this one
                    if not keyword_found:
                        keyword_found = True
                        first_keyword = keyword

                        # Extract bounding box coordinates
                        if len(bbox) == 4:
                            x_coords = [p[0] for p in bbox]
                            y_coords = [p[1] for p in bbox]
                            x_min, x_max = min(x_coords), max(x_coords)
                            y_min, y_max = min(y_coords), max(y_coords)

                            first_keyword_pos = {
                                "x_min": x_min,
                                "y_min": y_min,
                                "x_max": x_max,
                                "y_max": y_max,
                                "text": text,
                            }

                        # Stop after finding the first keyword
                    break

            # Break the outer loop if a keyword is found
            if keyword_found:
                break

        # Process the first keyword occurrence
        if first_keyword and first_keyword_pos:
            # Adjust bounding box
            x_min, y_min, x_max, y_max = self.adjust_bounding_box(
                first_keyword_pos["x_min"],
                first_keyword_pos["y_min"],
                first_keyword_pos["x_max"],
                first_keyword_pos["y_max"],
                first_keyword,
                image_np.shape,
            )

            # Ensure coordinates are integers and within image bounds
            height, width = image_np.shape[:2]
            x_min = int(max(0, x_min))
            y_min = int(max(0, y_min))
            x_max = int(min(width, x_max))
            y_max = int(min(height, y_max))

            # Expand bounding box to capture more context
            x_min = max(0, x_min - 100)
            y_min = max(0, y_min - 100)
            x_max = min(width, x_max + 100)
            y_max = min(height, y_max + 100)

            # Save the cropped image
            try:
                cropped = Image.fromarray(image_np[y_min:y_max, x_min:x_max])

                # Extract address ONLY from the cropped screenshot
                cropped_np = np.array(cropped)
                screenshot_results = self.reader.readtext(
                    cropped_np, detail=1, paragraph=False
                )

                # Improved address extraction
                address_candidates = []
                for result in screenshot_results:
                    text = result[1].strip()

                    # Enhanced address detection heuristics
                    if (
                        len(text) > 3  # Reduced minimum length
                        and not any(keyword in text for keyword in self.keywords)
                        and
                        # Expanded Japanese address pattern matching
                        (
                            re.search(
                                r"(^|\s)(福岡県|長崎県|鳥取県|島根県|[^\s]+県)([^\s]+市|[^\s]+町|[^\s]+村)",
                                text,
                            )
                            or re.search(r"\d+[-丁目番地号]+", text)
                            or re.search(r"[^\s]+[丁目番地号]\d+", text)
                        )
                    ):
                        # Filter out obvious non-address texts
                        if not re.search(
                            r"^[0-9\-]+$", text
                        ):  # Exclude pure number strings
                            address_candidates.append(text)

                # Clean and combine address candidates
                if address_candidates:
                    full_address = " ".join(address_candidates)
                    cleaned_address = self.clean_address(full_address)

                    return (
                        True,
                        {first_keyword: cleaned_address},
                    )

            except Exception as e:
                self.logger.error(f"Error Process Image OCR: {e}")
                return False, None

        # If no keyword found, or no address extracted
        return False, None

    def adjust_bounding_box(self, x_min, y_min, x_max, y_max, keyword, image_shape):
        """
        Adjust bounding box with custom offsets, ensuring it stays within image boundaries.
        """
        # Get adjustments for the specific keyword
        left_adj, top_adj, right_adj, bottom_adj = self.bbox_adjustments.get(
            keyword, (0, 0, 0, 0)
        )

        # Apply adjustments
        x_min = max(0, x_min + left_adj)
        y_min = max(0, y_min + top_adj)
        x_max = min(image_shape[1], x_max + right_adj)
        y_max = min(image_shape[0], y_max + bottom_adj)

        return x_min, y_min, x_max, y_max

    def extract_text_from_pdf(self, pdf_path) -> str | None:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                full_text = "\n".join(
                    page.extract_text() for page in pdf.pages if page.extract_text()
                )
            return full_text
        except Exception as e:
            self.logger.error(f"Error extracting text from {pdf_path}: {e}")
            return None

    def process_pdf(self, pdf_path: str) -> dict:
        """Process PDF File"""
        if not pdf_path.lower().endswith(".pdf"):
            self.logger.error(f"Required PDF file: {pdf_path}")
            return None
        self.logger.info(f"Processing: {pdf_path}")
        try:
            # Convert PDF to images and perform OCR'static/'
            pages = convert_from_path(
                poppler_path=self.poppler_path,
                pdf_path=pdf_path,
                dpi=self.dpi,
            )
            file_addresses: dict[str, str] = {}

            for _, page in enumerate(pages, 1):
                image_np = np.array(page)
                success, keyword_addresses = self.process_image_ocr(image_np)
                if success and keyword_addresses:
                    file_addresses.update(keyword_addresses)
                    break
            if not file_addresses:
                self.logger.info("No keywords found in the PDF.")
                return {}
            data: dict[str, str] = {}
            for keyword, address in file_addresses.items():
                data[keyword] = address
            return data
        except Exception as e:
            self.logger.error(f"Error processing {pdf_path}: {e}")
            return {}


__all__ = [PDFProcessor]
