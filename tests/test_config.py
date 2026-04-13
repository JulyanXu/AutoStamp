import json
import os
import pytest
from core.config import AppConfig


def test_default_config():
    config = AppConfig()
    assert config.stamp_path == ""
    assert config.stamp_scale == 100
    assert config.stamp_opacity == 100
    assert config.stamp_x == 50.0
    assert config.stamp_y == 50.0
    assert config.page_mode == "first"
    assert config.custom_pages == ""
    assert config.output_dir == ""


def test_save_and_load(tmp_path):
    path = tmp_path / "config.json"
    config = AppConfig()
    config.stamp_scale = 75
    config.stamp_opacity = 50
    config.stamp_x = 100.0
    config.stamp_y = 200.0
    config.page_mode = "custom"
    config.custom_pages = "1,3-5"
    config.save(str(path))

    loaded = AppConfig.load(str(path))
    assert loaded.stamp_scale == 75
    assert loaded.stamp_opacity == 50
    assert loaded.stamp_x == 100.0
    assert loaded.stamp_y == 200.0
    assert loaded.page_mode == "custom"
    assert loaded.custom_pages == "1,3-5"


def test_load_missing_file(tmp_path):
    path = tmp_path / "nonexistent.json"
    config = AppConfig.load(str(path))
    assert config.stamp_scale == 100


def test_parse_pages_first():
    config = AppConfig()
    config.page_mode = "first"
    assert config.get_page_indices(total_pages=5) == [0]


def test_parse_pages_last():
    config = AppConfig()
    config.page_mode = "last"
    assert config.get_page_indices(total_pages=5) == [4]


def test_parse_pages_all():
    config = AppConfig()
    config.page_mode = "all"
    assert config.get_page_indices(total_pages=5) == [0, 1, 2, 3, 4]


def test_parse_pages_custom():
    config = AppConfig()
    config.page_mode = "custom"
    config.custom_pages = "1,3-5"
    assert config.get_page_indices(total_pages=10) == [0, 2, 3, 4]


def test_parse_pages_custom_out_of_range():
    config = AppConfig()
    config.page_mode = "custom"
    config.custom_pages = "1,3,99"
    assert config.get_page_indices(total_pages=5) == [0, 2]
