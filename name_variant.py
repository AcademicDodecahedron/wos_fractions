from pathlib import Path
from typing import List

##

class NameVariants:
    def __init__(self, lines: List[str]):
        self.main_name = lines[0]

        lines = map(lambda x: x.lower(), lines) #type:ignore
        lines = sorted(lines, key=lambda x: len(x))
        self.lines = list(lines)

    def find_matching_pattern(self, name: str):
        for line in self.lines:
            if name.lower().find(line) != -1:
                return line
        return None

    @staticmethod
    def from_file(path: Path):
        with path.open('r', encoding='utf-8-sig') as file:
            lines = file.readlines()
            lines = map(lambda x: x.rstrip(), lines)
            return NameVariants(list(lines))

##

class NameVariantsDict:
    def __init__(self, name_variants: NameVariants):
        self.name_variants = name_variants
        self.name_dict = {}

    def find_matching_pattern(self, name: str):
        if name in self.name_dict:
            return self.name_dict[name]
        else:
            pattern = self.name_variants.find_matching_pattern(name)
            self.name_dict[name] = pattern
            return pattern

    def get_items(self):
        return self.name_dict.items()

    def get_main_name(self):
        return self.name_variants.main_name
