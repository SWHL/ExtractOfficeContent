# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import sys
import tempfile
from pathlib import Path

tests_dir = Path(__file__).resolve().parent
test_file_dir = tests_dir / "test_files"
root_dir = tests_dir.parent

sys.path.append(str(root_dir))

from extract_office_content import ExtractWord

word_extracter = ExtractWord()

word_path = test_file_dir / "word_example.docx"


def test_normal_input():
    res = word_extracter(word_path)

    assert res[:10] == "我与父亲不相见已二年"
    assert res[-2:] == " |"


def test_input_bytes():
    with open(word_path, "rb") as f:
        word_content = f.read()
    res = word_extracter(word_content)

    assert res[:10] == "我与父亲不相见已二年"
    assert res[-2:] == " |"


def test_extract_imgs():
    with tempfile.TemporaryDirectory() as tmp_dir:
        word_extracter(word_path, tmp_dir)

        img_list = list(Path(tmp_dir).iterdir())
        assert len(img_list) == 1
