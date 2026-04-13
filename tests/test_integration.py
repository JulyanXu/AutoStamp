import os
import fitz
from PIL import Image
import pytest
from core.config import AppConfig
from core.batch import BatchWorker


@pytest.fixture
def stamp_image(tmp_path):
    img = Image.new("RGBA", (80, 80), (0, 0, 0, 0))
    for x in range(80):
        for y in range(80):
            if (x - 40) ** 2 + (y - 40) ** 2 < 35 ** 2:
                img.putpixel((x, y), (255, 0, 0, 180))
    path = str(tmp_path / "stamp.png")
    img.save(path)
    return path


@pytest.fixture
def multi_page_pdf(tmp_path):
    doc = fitz.open()
    for i in range(5):
        page = doc.new_page(width=595, height=842)
        page.insert_text((100, 100), f"Page {i + 1}", fontsize=24)
    path = str(tmp_path / "multi.pdf")
    doc.save(path)
    doc.close()
    return path


def test_full_workflow_first_page(multi_page_pdf, stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=80,
        stamp_opacity=70,
        stamp_x=200.0,
        stamp_y=300.0,
        page_mode="first",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[multi_page_pdf], config=config)
    results = worker.process_all()
    assert results["success"] == 1
    assert results["failed"] == 0
    output_file = os.path.join(output_dir, "multi_已盖章.pdf")
    assert os.path.exists(output_file)
    doc = fitz.open(output_file)
    assert len(doc) == 5
    doc.close()


def test_full_workflow_custom_pages(multi_page_pdf, stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=100,
        stamp_opacity=100,
        stamp_x=50.0,
        stamp_y=50.0,
        page_mode="custom",
        custom_pages="1,3,5",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[multi_page_pdf], config=config)
    results = worker.process_all()
    assert results["success"] == 1
    output_file = os.path.join(output_dir, "multi_已盖章.pdf")
    assert os.path.exists(output_file)


def test_full_workflow_all_pages(multi_page_pdf, stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=150,
        stamp_opacity=50,
        stamp_x=0.0,
        stamp_y=0.0,
        page_mode="all",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[multi_page_pdf], config=config)
    results = worker.process_all()
    assert results["success"] == 1


def test_multiple_files(stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    pdfs = []
    for i in range(3):
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        page.insert_text((100, 100), f"Doc {i + 1}", fontsize=24)
        path = str(tmp_path / f"doc{i + 1}.pdf")
        doc.save(path)
        doc.close()
        pdfs.append(path)

    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=100,
        stamp_opacity=100,
        stamp_x=100.0,
        stamp_y=100.0,
        page_mode="first",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=pdfs, config=config)
    results = worker.process_all()
    assert results["success"] == 3
    assert results["failed"] == 0
    for i in range(3):
        assert os.path.exists(os.path.join(output_dir, f"doc{i + 1}_已盖章.pdf"))
