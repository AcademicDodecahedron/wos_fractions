from pathlib import Path
from typing import List

##

class NameVariants:
    def __init__(self, lines: List[str]):
        self.main_name = lines[0]

        lines = map(lambda x: x.lower(), lines)
        lines = sorted(lines, key=lambda x: len(x))
        self.lines = list(lines)

    def find_main_name(self, name: str):
        for line in self.lines:
            if name.lower().find(line) != -1:
                return self.main_name
        return None

    @staticmethod
    def from_file(path: Path):
        with path.open() as file:
            lines = file.readlines()
            lines = map(lambda x: x.removesuffix('\n'), lines)
            return NameVariants(list(lines))

##

class NameVariantsDict:
    def __init__(self, name_variants: NameVariants):
        self.name_variants = name_variants
        self.name_dict = {}

    def find_and_remember(self, name: str):
        if name in self.name_dict:
            return self.name_dict[name]
        else:
            main_name = self.name_variants.find_main_name(name)
            if main_name is None:
                main_name = 'NoN'

            self.name_dict[name] = main_name
            return main_name

    def get_items(self):
        return self.name_dict.items()
