import os
import fitz  # PyMuPDF
from PIL import Image
import pytest
from core.stamper import stamp_pdf


@pytest.fixture
def sample_pdf(tmp_path):
    """Create a simple 1-page PDF for testing."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4
    page.insert_text((100, 100), "Test Document", fontsize=24)
    path = str(tmp_path / "sample.pdf")
    doc.save(path)
    doc.close()
    return path


@pytest.fixture
def stamp_image(tmp_path):
    """Create a simple red circle PNG with transparency."""
    img = Image.new("RGBA", (100, 100), (0, 0, 0, 0))
    for x in range(100):
        for y in range(100):
            if (x - 50) ** 2 + (y - 50) ** 2 < 40 ** 2:
                img.putpixel((x, y), (255, 0, 0, 200))
    path = str(tmp_path / "stamp.png")
    img.save(path)
    return path


def test_stamp_pdf_creates_output(sample_pdf, stamp_image, tmp_path):
    output = str(tmp_path / "output.pdf")
    stamp_pdf(
        pdf_path=sample_pdf,
        stamp_path=stamp_image,
        output_path=output,
        x=100.0,
        y=100.0,
        scale=100,
        opacity=100,
        page_indices=[0],
    )
    assert os.path.exists(output)
    doc = fitz.open(output)
    assert len(doc) == 1
    doc.close()


def test_stamp_pdf_multiple_pages(tmp_path, stamp_image):
    doc = fitz.open()
    for i in range(3):
        page = doc.new_page(width=595, height=842)
        page.insert_text((100, 100), f"Page {i+1}", fontsize=24)
    pdf_path = str(tmp_path / "multi.pdf")
    doc.save(pdf_path)
    doc.close()

    output = str(tmp_path / "output.pdf")
    stamp_pdf(
        pdf_path=pdf_path,
        stamp_path=stamp_image,
        output_path=output,
        x=50.0,
        y=50.0,
        scale=100,
        opacity=80,
        page_indices=[0, 2],  # stamp pages 1 and 3
    )
    assert os.path.exists(output)
    doc = fitz.open(output)
    assert len(doc) == 3
    doc.close()


def test_stamp_pdf_with_scale(sample_pdf, stamp_image, tmp_path):
    """Stamp with scale=50 should produce a rect half the width of scale=100."""
    output_100 = str(tmp_path / "output_100.pdf")
    output_50 = str(tmp_path / "output_50.pdf")

    stamp_pdf(
        pdf_path=sample_pdf,
        stamp_path=stamp_image,
        output_path=output_100,
        x=0.0, y=0.0, scale=100, opacity=100, page_indices=[0],
    )
    stamp_pdf(
        pdf_path=sample_pdf,
        stamp_path=stamp_image,
        output_path=output_50,
        x=0.0, y=0.0, scale=50, opacity=100, page_indices=[0],
    )

    doc100 = fitz.open(output_100)
    doc50 = fitz.open(output_50)
    rects100 = doc100[0].get_image_rects(doc100.get_page_images(0)[0][0])
    rects50 = doc50[0].get_image_rects(doc50.get_page_images(0)[0][0])
    doc100.close()
    doc50.close()

    w100 = rects100[0].width
    w50 = rects50[0].width
    assert abs(w50 - w100 / 2) < 2.0, f"Expected scale=50 width to be ~{w100/2:.1f}, got {w50:.1f}"


def test_stamp_pdf_invalid_scale(sample_pdf, stamp_image, tmp_path):
    output = str(tmp_path / "output.pdf")
    with pytest.raises(ValueError):
        stamp_pdf(
            pdf_path=sample_pdf, stamp_path=stamp_image, output_path=output,
            x=0.0, y=0.0, scale=0, opacity=100, page_indices=[0],
        )
    with pytest.raises(ValueError):
        stamp_pdf(
            pdf_path=sample_pdf, stamp_path=stamp_image, output_path=output,
            x=0.0, y=0.0, scale=-10, opacity=100, page_indices=[0],
        )


def test_stamp_pdf_invalid_opacity(sample_pdf, stamp_image, tmp_path):
    output = str(tmp_path / "output.pdf")
    with pytest.raises(ValueError):
        stamp_pdf(
            pdf_path=sample_pdf, stamp_path=stamp_image, output_path=output,
            x=0.0, y=0.0, scale=100, opacity=101, page_indices=[0],
        )
    with pytest.raises(ValueError):
        stamp_pdf(
            pdf_path=sample_pdf, stamp_path=stamp_image, output_path=output,
            x=0.0, y=0.0, scale=100, opacity=-1, page_indices=[0],
        )
