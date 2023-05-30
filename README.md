## Extract Office Text
- 出于对Office系列文档的内容提取需求


### TODO
- [ ] 完善提取PPT内容组织方式
  - [ ] 提供是否保存提取内容到txt中选项
  - [ ] 提供是否单独提取图像到指定目录选项
  - [ ] 是否对提取的图像过OCR，并保存到txt中选项
- [ ] 增加excel内容提取，支持多种输格式（makdown,html）
- [ ] 增加word内容提取
- [ ] 提供单独指定格式的内容提取，例如:
    ```bash
    pip install extract_office_text[word]
    pip install extract_office_text[ppt]
    ```