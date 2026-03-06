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

        try:
            # Read page dimensions from the source doc in the main process
            src_doc = fitz.open(input_path)
            total_pages = len(src_doc)
            src_doc.close()

            effective_workers = min(self.workers, total_pages)
            self.logger.info(
                "Starting optimization: %s | %d pages | %d DPI | %d worker(s)",
                input_path.name, total_pages, self.dpi, effective_workers,
            )

            # Build task args — each worker receives the path string (picklable)
            tasks = [(str(input_path), page_num, self.dpi) for page_num in range(total_pages)]

            # Render pages in parallel; collect into a dict keyed by page_num
            rendered: dict[int, Tuple[bytes, float, float]] = {}

            if effective_workers == 1:
                # Sequential fallback — avoids process-spawn overhead for small PDFs
                for task in tasks:
                    page_num, img_bytes, w, h = _render_page(task)
                    rendered[page_num] = (img_bytes, w, h)
                    if (page_num + 1) % 10 == 0:
                        self.logger.info("Processed %d/%d pages...", page_num + 1, total_pages)
            else:
                with ProcessPoolExecutor(max_workers=effective_workers) as executor:
                    futures = {executor.submit(_render_page, task): task[1] for task in tasks}
                    completed = 0
                    for future in as_completed(futures):
                        page_num, img_bytes, w, h = future.result()
                        rendered[page_num] = (img_bytes, w, h)
                        completed += 1
                        if completed % 10 == 0:
                            self.logger.info("Rendered %d/%d pages...", completed, total_pages)

            # Assemble output PDF in correct page order (must be single-threaded)
            out_doc = fitz.open()
            for page_num in range(total_pages):
                img_bytes, w, h = rendered[page_num]
                out_page = out_doc.new_page(width=w, height=h)
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

        except Exception as e:
            self.logger.error("Failed to process document: %s", str(e), exc_info=True)
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