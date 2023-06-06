# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import tempfile
import sys
from pathlib import Path

tests_dir = Path(__file__).resolve().parent
test_file_dir = tests_dir / 'test_files'
root_dir = tests_dir.parent

sys.path.append(str(root_dir))

from extract_office_text.utils import read_txt
from extract_office_text import ExtractPPT

ppt_extracter = ExtractPPT()
ppt_path = test_file_dir / 'ppt_example.pptx'


def test_normal_input():
    res = ppt_extracter(ppt_path)

    assert len(res) == 26
    assert res[-1][:13] == 'www.ypppt.com'


def test_with_images():
    with tempfile.TemporaryDirectory() as tmp_dir:
        res = ppt_extracter(ppt_path, save_img_dir=tmp_dir)

        img_list = list(Path(tmp_dir).glob('*.*'))
        assert len(img_list) == 38


def test_with_txt():
    with tempfile.TemporaryDirectory() as tmp_dir:
        txt_path = Path(tmp_dir) / '1.txt'
        res = ppt_extracter(ppt_path, save_txt_path=txt_path)

        content = read_txt(txt_path)
        assert content[-1][:13] == '字体下载：http://w'
