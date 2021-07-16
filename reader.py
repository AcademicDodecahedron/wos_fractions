from pathlib import Path

##

def read(path: Path):
    with path.open(mode='r', encoding='utf-8-sig') as file:
        current_field = None
        current_value = None
        record = {}

        for line in file:
            line = line.rstrip()
            if line.startswith(' '):
                # Продолжение поля на новой строке
                current_value.append(line.lstrip())
            else:
                if current_field != None:
                    record[current_field] = current_value

                if line == 'ER':
                    # Конец записи
                    yield record
                    record = {}
                elif line != '' and line != 'EF':
                    # Новое поле
                    line_parts = line.split(' ', 1)
                    current_field = line_parts[0]
                    current_value = [line_parts[1]]
