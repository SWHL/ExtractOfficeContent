# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import sys
from pathlib import Path

tests_dir = Path(__file__).resolve().parent
test_file_dir = tests_dir / 'test_files'
root_dir = tests_dir.parent

sys.path.append(str(root_dir))

from extract_office_text import ExtractWord

word_extracter = ExtractWord()


def test_normal_input():
    word_path = test_file_dir / 'word_example.docx'
    res = word_extracter(word_path)

    assert len(res) == 316
    assert res[:10] == '我与父亲不相见已二年'
    assert res[-2:] == '提莫'
