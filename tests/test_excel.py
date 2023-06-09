# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import sys
import tempfile
from pathlib import Path

import pytest

tests_dir = Path(__file__).resolve().parent
test_file_dir = tests_dir / "test_files"
root_dir = tests_dir.parent

sys.path.append(str(root_dir))

from extract_office_content import ExtractExcel

excel_extracter = ExtractExcel()


@pytest.mark.parametrize(
    "out_format, sheet_len, gt",
    [("markdown", 2, "|    | 班级"), ("html", 2, "<table bo")],
)
def test_normal_input(out_format, sheet_len, gt):
    excel_path = test_file_dir / "excel_example.xlsx"
    res = excel_extracter(excel_path, out_format=out_format)

    assert len(res) == sheet_len
    assert res[0][:9] == gt


@pytest.mark.parametrize(
    "out_format, sheet_len, gt",
    [("markdown", 2, "|    | 班级"), ("html", 2, "<table bo")],
)
def test_input_bytes(out_format, sheet_len, gt):
    excel_path = test_file_dir / "excel_example.xlsx"
    with open(excel_path, "rb") as f:
        excel_content = f.read()
    res = excel_extracter(excel_content, out_format=out_format)

    assert len(res) == sheet_len
    assert res[0][:9] == gt


def test_with_images():
    excel_path = test_file_dir / "excel_with_image.xlsx"
    with tempfile.TemporaryDirectory() as tmp_dir:
        res = excel_extracter(excel_path, save_img_dir=tmp_dir)

        img_list = list(Path(tmp_dir).glob("*.*"))
        assert len(img_list) == 2
        assert res[0][:9] == "|    | 班级"


def test_without_images():
    excel_path = test_file_dir / "excel_example.xlsx"
    with tempfile.TemporaryDirectory() as tmp_dir:
        with pytest.warns(UserWarning, match="does not contain any images."):
            res = excel_extracter(excel_path, save_img_dir=tmp_dir)
