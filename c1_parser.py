# (c) E-kvadrat Consulting & Media, 2021

import re
from typing import Union
from collections import defaultdict

##

def add_aff_to_dict(authors, uni, c1_dict: defaultdict):
    if authors is not None:
        for author in set(authors.split('; ')):
            c1_dict[author].append(uni)

##

C1_REGEX = re.compile(r'^\[(.+)\] (.+)')
def parse(lines) -> Union[dict, str]:
    if len(lines) == 0:
        return {}

    # Нет указания автора
    if not lines[0].startswith('['):
        single_uni = lines[0]
        for line in lines[1:]:
            assert not line.startswith('['), f'Expected not authors in C1, got {line}'
            single_uni += ' ' + line
        return single_uni


    c1_dict = defaultdict(list)
    current_authors = None
    current_uni = ''

    for line in lines:
        if line.startswith('['):
            add_aff_to_dict(current_authors, current_uni, c1_dict)

            match = C1_REGEX.match(line)
            assert match, f'Failed to parse C1 line {line}'

            current_authors = match.group(1)
            current_uni = match.group(2)
        else:
            current_uni = ' '.join([current_uni, line])

    add_aff_to_dict(current_authors, current_uni, c1_dict)
    return dict(c1_dict)
