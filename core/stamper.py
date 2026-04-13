import fitz  # PyMuPDF
from PIL import Image
import io


def stamp_pdf(
    pdf_path: str,
    stamp_path: str,
    output_path: str,
    x: float,
    y: float,
    scale: int,
    opacity: int,
    page_indices: list[int],
) -> None:
    """
    Stamp a PNG image onto specified pages of a PDF.

    Args:
        pdf_path: Path to source PDF.
        stamp_path: Path to stamp PNG image (transparent background).
        output_path: Path to save the stamped PDF.
        x: X position in PDF points (from left edge).
        y: Y position in PDF points (from top edge).
        scale: Scale percentage (100 = original size).
        opacity: Opacity percentage (100 = fully opaque).
        page_indices: 0-based list of page indices to stamp.
    """
    # Load and scale the stamp image
    stamp_img = Image.open(stamp_path).convert("RGBA")
    if scale != 100:
        new_w = int(stamp_img.width * scale / 100)
        new_h = int(stamp_img.height * scale / 100)
        stamp_img = stamp_img.resize((new_w, new_h), Image.LANCZOS)

    # Apply opacity
    if opacity < 100:
        r, g, b, a = stamp_img.split()
        a = a.point(lambda p: int(p * opacity / 100))
        stamp_img = Image.merge("RGBA", (r, g, b, a))

    # Convert to PNG bytes for PyMuPDF
    buf = io.BytesIO()
    stamp_img.save(buf, format="PNG")
    stamp_bytes = buf.getvalue()

    stamp_w = stamp_img.width * 72 / 96  # pixels to PDF points (assuming 96 DPI)
    stamp_h = stamp_img.height * 72 / 96

    doc = fitz.open(pdf_path)
    for page_idx in page_indices:
        if page_idx < 0 or page_idx >= len(doc):
            continue
        page = doc[page_idx]
        rect = fitz.Rect(x, y, x + stamp_w, y + stamp_h)
        page.insert_image(rect, stream=stamp_bytes)

    doc.save(output_path)
    doc.close()
