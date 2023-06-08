# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
import sys
from pathlib import Path
from typing import List

import setuptools
from get_pypi_latest_version import GetPyPiLatestVersion


def get_readme() -> str:
    root_dir = Path(__file__).resolve().parent
    readme_path = str(root_dir / 'docs' / 'doc_whl.md')
    with open(readme_path, 'r', encoding='utf-8') as f:
        readme = f.read()
    return readme


def read_txt(txt_path: str) -> List:
    if not isinstance(txt_path, str):
        txt_path = str(txt_path)

    with open(txt_path, 'r', encoding='utf-8') as f:
        data = list(map(lambda x: x.rstrip('\n'), f))
    return data


MODULE_NAME = 'extract_office_content'

obtainer = GetPyPiLatestVersion()
try:
    latest_version = obtainer(MODULE_NAME)
except ValueError:
    latest_version = '0.0.1'
VERSION_NUM = obtainer.version_add_one(latest_version)

if len(sys.argv) > 2:
    match_str = ' '.join(sys.argv[2:])
    matched_versions = obtainer.extract_version(match_str)
    if matched_versions:
        VERSION_NUM = matched_versions
sys.argv = sys.argv[:2]

require = read_txt('requirements.txt')
setuptools.setup(
    name=MODULE_NAME,
    version=VERSION_NUM,
    platforms="Any",
    description="Tool for extracting content from office files.",
    long_description=get_readme(),
    long_description_content_type='text/markdown',
    author="SWHL",
    author_email="liekkaskono@163.com",
    url="https://github.com/SWHL/ExtractOfficeText.git",
    license='Apache-2.0',
    include_package_data=True,
    install_requires=require,
    packages=[MODULE_NAME],
    keywords=['extract,office,text'],
    classifiers=[
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
    ],
    python_requires='>=3.7,<=3.10',
    entry_points={
        'console_scripts': [f'{MODULE_NAME}={MODULE_NAME}.{MODULE_NAME}:main'],
    }
)
