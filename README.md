## Extract Office Text
- 出于对Office系列文档的内容提取需求


### TODO
- 完善提取PPT内容组织方式
  - [x] 提供是否保存提取内容到txt中选项
  - [x] 提供是否单独提取图像到指定目录选项
  - [ ] 是否对提取的图像过OCR，并保存到txt中选项
- [x] 增加excel内容提取，支持多种输格式（makdown,html）
- [ ] excel中图像的提取
- [ ] 增加word内容提取
- [ ] 提供单独指定格式的内容提取，例如:
    ```bash
    pip install extract_office_text[word]
    pip install extract_office_text[ppt]
    ```

### 提取PPT中文本和图像
```python
from pathlib import Path

from extract_office_text import ExtractPPTText


ppt_path = 'tests/test_files/ppt_example.pptx'

ppt_extracter = ExtractPPTText()

save_dir = 'outputs'
save_img_dir = Path(save_dir) / Path(ppt_path).stem

res = ppt_extracter(ppt_path,
                    is_save_img=True,
                    save_img_dir=str(save_img_dir),
                    is_save_to_txt=True,
                    save_txt_dir=save_dir)
print(res)
```

### 提取Excel中文本
```python
from extract_office_text import ExtractExcel

excel_extract = ExtractExcel()

excel_path = 'tests/test_files/excel_example.xlsx'

res  = excel_extract(excel_path, out_format='markdown')

print(res)
```

### 参考资料
- [Pandas读取excel合并单元格的正确姿势（openpyxl合并单元格拆分并填充内容）](https://blog.51cto.com/u_11466419/6100833)