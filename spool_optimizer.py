import argparse
import logging
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Tuple

import fitz


def _render_page(args: Tuple[str, int, int]) -> Tuple[int, bytes, float, float]:
    """Render a single PDF page to a grayscale JPEG in an isolated worker process.

    This function must live at module level so that multiprocessing can pickle it
    when spawning worker processes on all platforms (including Windows).

    Args:
        args: Tuple of (pdf_path_str, page_num, dpi).

    Returns:
        Tuple of (page_num, jpeg_bytes, page_width, page_height).
    """
    pdf_path_str, page_num, dpi = args
    doc = fitz.open(pdf_path_str)
    page = doc.load_page(page_num)
    pix = page.get_pixmap(dpi=dpi, alpha=False, colorspace=fitz.csGRAY)
    img_bytes = pix.tobytes("jpeg")
    width = page.rect.width
    height = page.rect.height
    doc.close()
    return page_num, img_bytes, width, height


class DocumentSpoolOptimizer:
    """Flattens and compresses PDF documents by rasterizing each page to grayscale JPEG
    images, reducing processing load on printer hardware and preventing memory overflow.

    Pages are rendered in parallel using a multiprocessing pool so that large PDFs
    benefit from all available CPU cores, then reassembled in the correct order.
    """

    def __init__(self, dpi: int = 100, workers: int = 0):
        """Initialize the optimizer.

        Args:
            dpi: Dots per inch for rasterization. Must be between 72 and 300.
                 Lower values produce smaller files; higher values retain more detail.
            workers: Number of worker processes for parallel rendering.
                     0 (default) means use all logical CPU cores.
                     1 disables multiprocessing and runs sequentially.
        """
        self.dpi = dpi
        self.workers = workers if workers > 0 else (os.cpu_count() or 1)
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
        """Flatten and compress a PDF document using parallel page rendering.

        Pages are rasterized concurrently across worker processes then reassembled
        in the correct order into a single compressed output PDF.

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
            # Attempt to open the PDF and perform validation checks
            src_doc = fitz.open(input_path)
            
            # Check if the PDF is encrypted/password-protected
            if src_doc.is_encrypted:
                self.logger.error("PDF is password-protected: %s", input_path.name)
                src_doc.close()
                return False
            
            # Check if the PDF has pages (detect corrupted/empty files)
            total_pages = len(src_doc)
            if total_pages == 0:
                self.logger.error("PDF has no pages or is corrupted: %s", input_path.name)
                src_doc.close()
                return False
            
            self.logger.info("Total pages to process: %d", total_pages)
            
            # Validate that at least the first page can be loaded and rendered
            try:
                first_page = src_doc.load_page(0)
                first_page.get_pixmap(dpi=self.dpi, alpha=False, colorspace=fitz.csGRAY)
                self.logger.info("Validation passed: First page rendered successfully")
            except Exception as e:
                self.logger.error("Content stream corruption - cannot render pages: %s - %s", input_path.name, str(e))
                src_doc.close()
                return False
            
            src_doc.close()
            
            # Use multiprocessing for parallel page rendering
            pdf_path_str = str(input_path.absolute())
            tasks = [(pdf_path_str, page_num, self.dpi) for page_num in range(total_pages)]
            
            rendered_pages = {}
            
            if self.workers == 1:
                # Sequential processing
                for page_num in range(total_pages):
                    page_num_result, img_bytes, width, height = _render_page((pdf_path_str, page_num, self.dpi))
                    rendered_pages[page_num_result] = (img_bytes, width, height)
                    if (page_num + 1) % 10 == 0:
                        self.logger.info("Processed %d/%d pages...", page_num + 1, total_pages)
            else:
                # Parallel processing
                with ProcessPoolExecutor(max_workers=self.workers) as executor:
                    futures = {executor.submit(_render_page, task): task[1] for task in tasks}
                    
                    for future in as_completed(futures):
                        page_num, img_bytes, width, height = future.result()
                        rendered_pages[page_num] = (img_bytes, width, height)
                        
                        if len(rendered_pages) % 10 == 0:
                            self.logger.info("Processed %d/%d pages...", len(rendered_pages), total_pages)
            
            # Assemble rendered pages into output PDF
            out_doc = fitz.open()
            for page_num in range(total_pages):
                img_bytes, width, height = rendered_pages[page_num]
                out_page = out_doc.new_page(width=width, height=height)
                out_page.insert_image(out_page.rect, stream=img_bytes)

            out_doc.save(output_path, garbage=4, deflate=True, clean=True)
            out_doc.close()

            elapsed_time = time.time() - start_time
            self.logger.info(
                "Optimization complete. Saved to: %s. Time taken: %.2fs",
                output_path.name, elapsed_time,
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
    parser.add_argument(
        "--workers", type=int, default=0,
        help="Worker processes for parallel rendering (default: 0 = all CPU cores).",
    )

    args = parser.parse_args()

    if not args.input.exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if args.dpi < 72 or args.dpi > 300:
        print(f"Error: DPI must be between 72 and 300, got {args.dpi}", file=sys.stderr)
        sys.exit(1)

    if args.workers < 0:
        print(f"Error: --workers must be 0 or a positive integer, got {args.workers}", file=sys.stderr)
        sys.exit(1)

    optimizer = DocumentSpoolOptimizer(dpi=args.dpi, workers=args.workers)
    success = optimizer.process_document(args.input, args.output)

    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()