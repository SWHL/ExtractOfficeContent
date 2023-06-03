# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import sys
from pathlib import Path

import pytest

tests_dir = Path(__file__).resolve().parent
test_file_dir = tests_dir / 'test_files'
root_dir = tests_dir.parent

sys.path.append(str(root_dir))

from extract_office_text import ExtractExcel

excel_extracter = ExtractExcel()


@pytest.mark.parametrize(
    'out_format, sheet_len, gt',
    [
        ('markdown', 2, '|    | 班级'),
        ('html', 2, '<table bo')
    ]
)
def test_normal_input(out_format, sheet_len, gt):
    excel_path = test_file_dir / 'excel_example.xlsx'
    res = excel_extracter(excel_path, out_format=out_format)

    assert len(res) == sheet_len
    assert res[0][:9] == gt
