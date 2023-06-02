# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import sys
from pathlib import Path

tests_dir = Path(__file__).resolve().parent
test_file_dir = tests_dir / 'test_files'
root_dir = tests_dir.parent

sys.path.append(str(root_dir))

from extract_office_text import ExtractPPTText

ppt_extracter = ExtractPPTText()


def test_normal_input():
    ppt_path = test_file_dir / 'ppt_example.pptx'
    res = ppt_extracter(ppt_path)

    assert len(res) == 26
    assert res[-1][:13] == 'www.ypppt.com'
