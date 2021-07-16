import re

C1_REGEX = re.compile(r'^\[(.+)\] (.+)')
def parse(lines):
    c1_dict = {}
    for line in lines:
        match = C1_REGEX.match(line)
        if match is None:
            print(f"Unexpected line in C1: {line}")

        authors = match.group(1)
        uni = match.group(2)

        for author in authors.split('; '):
            if author in c1_dict:
                c1_dict[author].append(uni)
            else:
                c1_dict[author] = [uni]
    return c1_dict
