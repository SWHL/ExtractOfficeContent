# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import tempfile
import sys
from pathlib import Path

tests_dir = Path(__file__).resolve().parent
test_file_dir = tests_dir / "test_files"
root_dir = tests_dir.parent

sys.path.append(str(root_dir))

from extract_office_content import ExtractPPT

ppt_extracter = ExtractPPT()
ppt_path = test_file_dir / "ppt_example.pptx"


def test_normal_input():
    res = ppt_extracter(ppt_path)

    assert len(res) == 1
    assert res[-1][:3] == "| 0"


def test_input_bytes():
    with open(ppt_path, "rb") as f:
        ppt_content = f.read()
    res = ppt_extracter(ppt_content)

    assert len(res) == 1
    assert res[-1][:3] == "| 0"


def test_with_images():
    with tempfile.TemporaryDirectory() as tmp_dir:
        res = ppt_extracter(ppt_path, save_img_dir=tmp_dir)

        img_list = list(Path(tmp_dir).glob("*.*"))
        assert len(img_list) == 2
