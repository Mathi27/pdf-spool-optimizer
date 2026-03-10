import argparse
import logging
import sys
import time
from pathlib import Path
import fitz   


class DocumentSpoolOptimizer:
    """Flattens and compresses PDF documents by rasterizing each page to grayscale JPEG
    images, reducing processing load on printer hardware and preventing memory overflow."""

    def __init__(self, dpi: int = 100):
        """Initialize the optimizer with a target rasterization DPI.

        Args:
            dpi: Dots per inch for rasterization. Must be between 72 and 300.
                 Lower values produce smaller files; higher values retain more detail.
        """
        self.dpi = dpi
        self.logger = self._setup_logger()

    @staticmethod
    def _setup_logger() -> logging.Logger:
        """Create and configure a stdout logger for this module.

        Returns:
            A configured Logger instance.
        """
        logger = logging.getLogger(__name__)
        if not logger.handlers:
            logger.setLevel(logging.INFO)
            handler = logging.StreamHandler(sys.stdout)
            formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s'
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        return logger

    def process_document(self, input_path: Path, output_path: Path) -> bool:
        """Flatten and compress a PDF document.

        Each page is rasterized to a grayscale JPEG at the configured DPI and
        reassembled into a new PDF with maximum compression.

        Args:
            input_path: Path to the source PDF file.
            output_path: Destination path for the optimized PDF.

        Returns:
            True on success, False if the input file is missing or any error occurs.
        """
        if not input_path.exists():
            self.logger.error("Input file does not exist: %s", input_path)
            return False

        start_time = time.time()
        self.logger.info("Starting optimization for: %s at %d DPI (Grayscale)", input_path.name, self.dpi)

        try:
            # Attempt to open the PDF
            src_doc = fitz.open(input_path)
            
            # Check if the PDF is encrypted/password-protected
            if src_doc.is_encrypted:
                self.logger.error("PDF is password-protected: %s", input_path.name)
                src_doc.close()
                return False
            
            # Check if the PDF has pages (detect corrupted/empty files)
            if len(src_doc) == 0:
                self.logger.error("PDF has no pages or is corrupted: %s", input_path.name)
                src_doc.close()
                return False
            
            total_pages = len(src_doc)
            
            # PRIORITY 1: Verify page tree integrity by attempting to load first page
            try:
                first_page = src_doc.load_page(0)
            except Exception as e:
                self.logger.error("Page tree corruption - cannot load first page: %s - %s", 
                                input_path.name, str(e))
                src_doc.close()
                return False
            
            # Check for extreme page dimensions that could cause memory issues
            page_width = first_page.rect.width
            page_height = first_page.rect.height
            max_dimension = 14400  # 200 inches at 72 DPI, reasonable maximum
            
            if page_width > max_dimension or page_height > max_dimension or page_width <= 0 or page_height <= 0:
                self.logger.error("Invalid page dimensions (%dx%d) in PDF: %s", 
                                int(page_width), int(page_height), input_path.name)
                src_doc.close()
                return False
            
            # PRIORITY 2: Verify content stream integrity by attempting to render first page
            try:
                test_pix = first_page.get_pixmap(dpi=self.dpi, alpha=False, colorspace=fitz.csGRAY)
                test_pix = None  # Release memory
            except Exception as e:
                self.logger.error("Content stream corruption - cannot render first page: %s - %s", 
                                input_path.name, str(e))
                src_doc.close()
                return False
            
            self.logger.info("Total pages to process: %d", total_pages)
            
            out_doc = fitz.open()

            for page_num in range(total_pages):
                try:
                    page = src_doc.load_page(page_num)
                except Exception as e:
                    self.logger.error("Failed to load page %d/%d: %s - %s", 
                                    page_num + 1, total_pages, input_path.name, str(e))
                    src_doc.close()
                    out_doc.close()
                    return False
                
                try:
                    pix = page.get_pixmap(dpi=self.dpi, alpha=False, colorspace=fitz.csGRAY)
                except Exception as e:
                    self.logger.error("Failed to render page %d/%d - content stream may be corrupted: %s", 
                                    page_num + 1, total_pages, str(e))
                    src_doc.close()
                    out_doc.close()
                    return False
 
                img_bytes = pix.tobytes("jpeg")
                
                out_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
                out_page.insert_image(out_page.rect, stream=img_bytes)
                
                if (page_num + 1) % 10 == 0:
                    self.logger.info("Processed %d/%d pages...", page_num + 1, total_pages)

          
            out_doc.save(
                output_path,
                garbage=4,
                deflate=True,
                clean=True
            )
            
            src_doc.close()
            out_doc.close()

            elapsed_time = time.time() - start_time
            self.logger.info(
                "Optimization complete. Saved to: %s. Time taken: %.2fs", 
                output_path.name, elapsed_time
            )
            
            self._log_compression_ratio(input_path, output_path)
            return True

        except fitz.FileDataError as e:
            error_msg = str(e).lower()
            if "xref" in error_msg or "cross-reference" in error_msg:
                self.logger.error("Cross-reference table corruption in PDF: %s - %s", input_path.name, str(e))
            elif "header" in error_msg or "pdf" in error_msg[:20]:
                self.logger.error("Invalid PDF header or structure: %s - %s", input_path.name, str(e))
            elif "compression" in error_msg or "flate" in error_msg or "deflate" in error_msg:
                self.logger.error("Compression/decompression error in PDF: %s - %s", input_path.name, str(e))
            else:
                self.logger.error("Corrupted or invalid PDF file: %s - %s", input_path.name, str(e))
            return False
        except fitz.FileNotFoundError as e:
            self.logger.error("PDF file not found: %s - %s", input_path.name, str(e))
            return False
        except RuntimeError as e:
            error_msg = str(e).lower()
            if "password" in error_msg or "encrypted" in error_msg:
                self.logger.error("PDF is password-protected: %s", input_path.name)
            elif "damaged" in error_msg or "corrupt" in error_msg:
                self.logger.error("PDF file is corrupted: %s", input_path.name)
            else:
                self.logger.error("Runtime error processing PDF: %s - %s", input_path.name, str(e))
            return False
        except MemoryError:
            self.logger.error("Insufficient memory to process PDF: %s", input_path.name)
            return False
        except Exception as e:
            self.logger.error("Failed to process document: %s - %s", input_path.name, str(e), exc_info=True)
            return False

    def _log_compression_ratio(self, original: Path, optimized: Path) -> None:
        """Log the size comparison between the original and optimized PDF files.

        Args:
            original: Path to the original input PDF.
            optimized: Path to the newly written output PDF.
        """
        orig_size_mb = original.stat().st_size / (1024 * 1024)
        opt_size_mb = optimized.stat().st_size / (1024 * 1024)
        
        self.logger.info("Original Size: %.2f MB", orig_size_mb)
        self.logger.info("Optimized Size: %.2f MB", opt_size_mb)
        
        if orig_size_mb > 0:
            ratio = (opt_size_mb / orig_size_mb) * 100
            self.logger.info("Output is %.2f%% of original size.", ratio)


def main() -> None:
    """CLI entry point: parse arguments and run the document optimizer."""
    parser = argparse.ArgumentParser(
        description="Flatten and compress PDF notes to optimize print spooling."
    )
    parser.add_argument("-i", "--input", required=True, type=Path, help="Path to input PDF file.")
    parser.add_argument("-o", "--output", required=True, type=Path, help="Path for output PDF file.")
    parser.add_argument("--dpi", type=int, default=100, help="Rasterization DPI (default: 100).")
    
    args = parser.parse_args()

    if not args.input.exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if args.dpi < 72 or args.dpi > 300:
        print(f"Error: DPI must be between 72 and 300, got {args.dpi}", file=sys.stderr)
        sys.exit(1)

    optimizer = DocumentSpoolOptimizer(dpi=args.dpi)
    success = optimizer.process_document(args.input, args.output)

    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()