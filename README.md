## Extract Office Content
<p>
    <a href=""><img src="https://img.shields.io/badge/Python->=3.7,<=3.10-aff.svg"></a>
    <a href=""><img src="https://img.shields.io/badge/OS-Linux%2C%20Win%2C%20Mac-pink.svg"></a>
    <a href="https://pypi.org/project/extract_office_content/"><img alt="PyPI" src="https://img.shields.io/pypi/v/extract_office_content"></a>
    <a href="https://pepy.tech/project/extract_office_content"><img src="https://static.pepy.tech/personalized-badge/extract_office_content?period=total&units=abbreviation&left_color=grey&right_color=blue&left_text=Downloads"></a>
</p>

- 出于对Office系列文档的内容提取需求


### TODO
- 完善提取PPT内容组织方式
  - [x] 提供是否保存提取内容到txt中选项
  - [x] 提供是否单独提取图像到指定目录选项
  - [ ] 是否对提取的图像过OCR，并保存到txt中选项
- [x] 增加excel内容提取，支持多种输格式（makdown,html）
- [x] excel中图像的提取
- [x] 增加word内容提取
  - [x] 支持提取表格
  - [x] 支持提取图像
- [ ] 提供单独指定格式的内容提取，例如:
    ```bash
    pip install extract_office_text[word]
    pip install extract_office_text[ppt]
    ```

### 提取PPT内容
```python
from pathlib import Path

from extract_office_text import ExtractPPT


ppt_path = 'tests/test_files/ppt_example.pptx'

ppt_extracter = ExtractPPT()

save_dir = 'outputs'
save_img_dir = Path(save_dir) / Path(ppt_path).stem
res = ppt_extracter(ppt_path,
                    save_img_dir=str(save_img_dir),
                    save_txt_path=Path(save_dir) / '1.txt')
print(res)
```

### 提取Excel内容
```python
from extract_office_text import ExtractExcel

excel_extract = ExtractExcel()
excel_path = 'tests/test_files/excel_example.xlsx'
res  = excel_extract(excel_path, out_format='markdown', save_img_dir='1')
print(res)
```

### 提取Word内容
```python
from extract_office_text import ExtractWord


word_extract = ExtractWord()

word_path = 'tests/test_files/word_example.docx'
text = word_extract(word_path, "outputs/word")
print(text)
```

### 参考资料
- [Pandas读取excel合并单元格的正确姿势（openpyxl合并单元格拆分并填充内容）](https://blog.51cto.com/u_11466419/6100833)
- [python-docx2txt](https://github.com/ankushshah89/python-docx2txt)