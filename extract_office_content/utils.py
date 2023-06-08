# -*- encoding: utf-8 -*-
# @Author: SWHL
# @Contact: liekkaskono@163.com
from pathlib import Path
from typing import List, Union


def mkdir(dir_path):
    Path(dir_path).mkdir(parents=True, exist_ok=True)


def read_txt(txt_path: str) -> List:
    if not isinstance(txt_path, str):
        txt_path = str(txt_path)

    with open(txt_path, 'r', encoding='utf-8') as f:
        data = list(map(lambda x: x.rstrip('\n'), f))
    return data


def write_txt(save_path: Union[str, Path],
              content: list, mode: str = 'w'):
    if not isinstance(save_path, str):
        save_path = str(save_path)

    if not isinstance(content, list):
        content = [content]

    with open(save_path, mode, encoding='utf-8') as f:
        for value in content:
            f.write(f'{value}\n')


def is_contain(sentence: str, key_words: Union[str, List],) -> bool:
    """sentences中是否包含key_words中任意一个"""
    return any(i in sentence for i in key_words)
