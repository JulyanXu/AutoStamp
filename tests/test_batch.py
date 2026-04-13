import os
import fitz
from PIL import Image
import pytest
from core.batch import BatchWorker
from core.config import AppConfig


@pytest.fixture
def stamp_image(tmp_path):
    img = Image.new("RGBA", (50, 50), (255, 0, 0, 200))
    path = str(tmp_path / "stamp.png")
    img.save(path)
    return path


@pytest.fixture
def sample_pdf(tmp_path):
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((100, 100), "Test", fontsize=24)
    path = str(tmp_path / "test.pdf")
    doc.save(path)
    doc.close()
    return path


def test_batch_process_single_pdf(sample_pdf, stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=100,
        stamp_opacity=100,
        stamp_x=10.0,
        stamp_y=10.0,
        page_mode="first",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[sample_pdf], config=config)
    results = worker.process_all()
    assert results["success"] == 1
    assert results["failed"] == 0
    assert os.path.exists(os.path.join(output_dir, "test_已盖章.pdf"))


def test_batch_skips_bad_file(stamp_image, tmp_path):
    bad_file = str(tmp_path / "bad.pdf")
    with open(bad_file, "w") as f:
        f.write("not a pdf")
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=100,
        stamp_opacity=100,
        stamp_x=10.0,
        stamp_y=10.0,
        page_mode="first",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[bad_file], config=config)
    results = worker.process_all()
    assert results["failed"] == 1
    assert len(results["errors"]) == 1
