# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from pathlib import Path
import zipfile

excel_path = Path('tests/test_files/excel_with_image.xlsx')
zip_excel_path = excel_path.with_name(f'{excel_path.stem}.zip')

excel_path.rename(zip_excel_path)

zf = zipfile.ZipFile(zip_excel_path)
zf.extractall(excel_path.parent / excel_path.stem)
zf.close()

zip_excel_path.rename(excel_path)
