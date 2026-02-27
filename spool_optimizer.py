"""
Optimized for high-speed print spooling of standard text documents and exam notes.
Implements aggressive DPI reduction and single-channel grayscale rasterization.
"""

import argparse
import logging
import sys
import time
from pathlib import Path
import fitz  # PyMuPDF


class DocumentSpoolOptimizer:
    """
    Handles the aggressive rasterization and compression of PDF documents.
    """

    def __init__(self, dpi: int = 100):
        """
        Initializes the optimizer with specified output parameters.

        Args:
            dpi (int): The resolution for rasterization. Default is 100 for basic text notes.
        """
        self.dpi = dpi
        self.logger = self._setup_logger()

    @staticmethod
    def _setup_logger() -> logging.Logger:
        """Configures the module logger."""
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
        """
        Flattens and compresses the input PDF document into grayscale.
        """
        if not input_path.exists():
            self.logger.error("Input file does not exist: %s", input_path)
            return False

        start_time = time.time()
        self.logger.info("Starting optimization for: %s at %d DPI (Grayscale)", input_path.name, self.dpi)

        try:
            src_doc = fitz.open(input_path)
            out_doc = fitz.open()

            total_pages = len(src_doc)
            self.logger.info("Total pages to process: %d", total_pages)

            for page_num in range(total_pages):
                page = src_doc.load_page(page_num)
                pix = page.get_pixmap(dpi=self.dpi, alpha=False, colorspace=fitz.csGRAY)
 
                img_bytes = pix.tobytes("jpeg")
                
                out_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
                out_page.insert_image(out_page.rect, stream=img_bytes)
                
                if (page_num + 1) % 10 == 0:
                    self.logger.info("Processed %d/%d pages...", page_num + 1, total_pages)

            # Apply Level-4 garbage collection and deflation
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

        except Exception as e:
            self.logger.error("Failed to process document: %s", str(e), exc_info=True)
            return False

    def _log_compression_ratio(self, original: Path, optimized: Path) -> None:
        """Calculates and logs the file size changes."""
        orig_size_mb = original.stat().st_size / (1024 * 1024)
        opt_size_mb = optimized.stat().st_size / (1024 * 1024)
        
        self.logger.info("Original Size: %.2f MB", orig_size_mb)
        self.logger.info("Optimized Size: %.2f MB", opt_size_mb)
        
        if orig_size_mb > 0:
            ratio = (opt_size_mb / orig_size_mb) * 100
            self.logger.info("Output is %.2f%% of original size.", ratio)


def main() -> None:
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Flatten and compress PDF notes to optimize print spooling."
    )
    parser.add_argument("-i", "--input", required=True, type=Path, help="Path to input PDF file.")
    parser.add_argument("-o", "--output", required=True, type=Path, help="Path for output PDF file.")
    parser.add_argument("--dpi", type=int, default=100, help="Rasterization DPI (default: 100).")
    
    args = parser.parse_args()

    optimizer = DocumentSpoolOptimizer(dpi=args.dpi)
    success = optimizer.process_document(args.input, args.output)
    
    if not success:
        sys.exit(1)


if __name__ == "__main__":
    main()