## Extract Office Content
<p>
    <a href="https://swhl-extractofficecontentdemo.hf.space" target="_blank"><img src="https://img.shields.io/badge/%F0%9F%A4%97-Online Demo-blue"></a>
    <a href=""><img src="https://img.shields.io/badge/Python->=3.6,<3.12-aff.svg"></a>
    <a href=""><img src="https://img.shields.io/badge/OS-Linux%2C%20Win%2C%20Mac-pink.svg"></a>
    <a href="https://pypi.org/project/extract_office_content/"><img alt="PyPI" src="https://img.shields.io/pypi/v/extract_office_content"></a>
    <a href="https://pepy.tech/project/extract_office_content"><img src="https://static.pepy.tech/personalized-badge/extract_office_content?period=total&units=abbreviation&left_color=grey&right_color=blue&left_text=Downloads"></a>
    <a href="https://semver.org/"><img alt="SemVer2.0" src="https://img.shields.io/badge/SemVer-2.0-brightgreen"></a>
    <a href="https://github.com/psf/black"><img src="https://img.shields.io/badge/code%20style-black-000000.svg"></a>
</p>

### 目前已知问题
- 提取PPT:
  - 提取ppt中的内容时，会丢失带有公式的文本框
  - 提取的表格格式不全
  - PPT中的表格会提取为对应的excel文件，是否有更好的方式？
- 提取Word:
  - 表格位置不能与原文中一一对应

### Use
1. Install`extract_office_content`
   ```bash
   $ pip install extract_office_content
   ```
2. Run by CLI.
    - Extract All office file's content.
        ```bash
        $ extract_office_content -h
        usage: extract_office_content [-h] [-img_dir SAVE_IMG_DIR] file_path

        positional arguments:
        file_path

        optional arguments:
        -h, --help            show this help message and exit
        -img_dir SAVE_IMG_DIR, --save_img_dir SAVE_IMG_DIR

        $ extract_office_content tests/test_files
        ```
    - Extract Word.
        ```bash
        $ extract_word -h
        usage: extract_word [-h] [-img_dir SAVE_IMG_DIR] word_path

        positional arguments:
        word_path

        optional arguments:
        -h, --help            show this help message and exit
        -img_dir SAVE_IMG_DIR, --save_img_dir SAVE_IMG_DIR

        $ extract_word tests/test_files/word_example.docx
        ```
    - Extract PPT.
        ```bash
        $ extract_ppt -h
        usage: extract_ppt [-h] [-img_dir SAVE_IMG_DIR] ppt_path

        positional arguments:
        ppt_path

        optional arguments:
        -h, --help            show this help message and exit
        -img_dir SAVE_IMG_DIR, --save_img_dir SAVE_IMG_DIR

        $ extract_ppt tests/test_files/ppt_example.pptx
        ```
    - Extract Excel.
        ```bash
        $ extract_excel -h
        usage: extract_excel [-h] [-f {markdown,html,latex,string}] [-o SAVE_IMG_DIR]
                            excel_path

        positional arguments:
        excel_path

        optional arguments:
        -h, --help            show this help message and exit
        -f {markdown,html,latex,string}, --output_format {markdown,html,latex,string}
        -o SAVE_IMG_DIR, --save_img_dir SAVE_IMG_DIR

        $ extract_excel tests/test_files/excel_example.xlsx
        ```
3. Run by python script.
    - Extract All.
        ```python
        from pathlib import Path

        from extract_office_content import ExtractOfficeContent

        extracter = ExtractOfficeContent()
        file_list = list(Path('tests/test_files').iterdir())

        for file_path in file_list:
            res = extracter(file_path)
            print(res)
        ```
    - Extract Word.
        ```python
        from extract_office_content import ExtractWord

        word_extract = ExtractWord()
        word_path = 'tests/test_files/word_example.docx'
        text = word_extract(word_path, "outputs/word")

        # or bytes
        with open(word_path, 'rb') as f:
            word_content = f.read()
        text = word_extract(word_content, "outputs/word")
        print(text)
        ```
    - Extract PPT.
        ```python
        from pathlib import Path

        from extract_office_content import ExtractPPT

        ppt_extracter = ExtractPPT()

        ppt_path = 'tests/test_files/ppt_example.pptx'
        save_dir = 'outputs'
        save_img_dir = Path(save_dir) / Path(ppt_path).stem
        res = ppt_extracter(ppt_path, save_img_dir=str(save_img_dir))

        # or bytes
        with open(ppt_path, 'rb') as f:
            ppt_content = f.read()
        res = ppt_extracter(ppt_content, save_img_dir=str(save_img_dir))
        print(res)
        ```
    - Extract Excel.
        ```python
        from extract_office_content import ExtractExcel

        excel_extract = ExtractExcel()

        excel_path = 'tests/test_files/excel_with_image.xlsx'
        res  = excel_extract(excel_path, out_format='markdown', save_img_dir='1')

        # or bytes
        with open(excel_path, 'rb') as f:
            excel_content = f.read()
        res  = excel_extract(excel_content, out_format='markdown', save_img_dir='1')
        print(res)
        ```


### 更新日志
- 2023-07-02 v0.0.6 update:
  - 统一提取word接口返回值为List，与其他统一
- 2023-06-17 v0.0.4 update:
  - 支持`file-like object`输入

### Reference
- [Pandas读取excel合并单元格的正确姿势（openpyxl合并单元格拆分并填充内容）](https://blog.51cto.com/u_11466419/6100833)
- [python-docx2txt](https://github.com/ankushshah89/python-docx2txt)
