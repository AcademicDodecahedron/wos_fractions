# (c) E-kvadrat Consulting & Media, 2021

import re
from typing import Union, NamedTuple, Optional
from collections import defaultdict

##


# Один блок вида [Автор; Автор] Университет, либо просто Университет
# может занимать несколько строк
class C1Block(NamedTuple):
    authors: Optional[set]
    uni: str
    length: int


C1_REGEX = re.compile(r"^\[(.+)\] (.+)")


def parse_next_block(lines):
    authors_text = None
    uni_text = None
    length = 1

    # Проверяем есть ли авторы на первой строке
    if lines[0].startswith("["):
        match = C1_REGEX.match(lines[0])
        assert match, f"Unexpected line in C1: {lines[0]}"

        authors_text = match.group(1)
        uni_text = match.group(2)
    else:
        uni_text = lines[0]

    # Остальные строки до начала нового блока
    # присоединяем через пробел
    for line in lines[1:]:
        if line.startswith("["):
            break
        uni_text += " " + line
        length += 1

    authors_split = set(authors_text.split("; ")) if authors_text is not None else None

    return C1Block(authors=authors_split, uni=uni_text, length=length)


##


def add_block_to_dict(block: C1Block, c1_dict: defaultdict):
    if block.authors:
        for author in block.authors:
            c1_dict[author].append(block.uni)


##


def parse(lines) -> Union[dict, str]:
    if len(lines) == 0:
        return {}

    c1_dict = defaultdict(list)

    first_block = parse_next_block(lines)
    if first_block.authors is None:
        # Если в первом блоке нет авторов, и это единственный блок - вернуть строку с названием университета
        # Если в первом блоке нет авторов, но есть в последующих блоках - игнорировать первый блок
        if first_block.length == len(lines):
            return first_block.uni
    else:
        add_block_to_dict(first_block, c1_dict)

    current_line = first_block.length
    while current_line < len(lines):
        next_block = parse_next_block(lines[current_line:])
        add_block_to_dict(next_block, c1_dict)
        current_line += next_block.length

    return dict(c1_dict)
