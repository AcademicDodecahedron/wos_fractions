# (c) E-kvadrat Consulting & Media, 2021

import re

##

def add_aff_to_dict(authors: str, uni, c1_dict):
    for author in set(authors.split('; ')):
        if author in c1_dict:
            c1_dict[author].append(uni)
        else:
            c1_dict[author] = [uni]

##

C1_REGEX = re.compile(r'^\[(.+)\] (.+)')
def parse(lines):
    c1_dict = {}
    current_authors = None
    current_uni = ''

    for line in lines:
        if line.startswith('['):
            if current_authors != None:
                add_aff_to_dict(current_authors, current_uni, c1_dict)

            match = C1_REGEX.match(line)
            current_authors = match.group(1)
            current_uni = match.group(2)
        else:
            current_uni = ' '.join([current_uni, line])

    if current_authors != None:
        add_aff_to_dict(current_authors, current_uni, c1_dict)
    return c1_dict
