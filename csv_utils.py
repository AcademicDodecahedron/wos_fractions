# (c) E-kvadrat Consulting & Media, 2021

from pathlib import Path
from io import StringIO
import csv

##

def read_csv_body(path: Path):
    body_str = ''
    with path.open(mode='r', encoding='utf-8-sig') as file:
        for line in file:
            if line == '\n':
                break
            body_str += line

    return csv.DictReader(StringIO(body_str))
