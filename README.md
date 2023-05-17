## Скрипт для фракционного расчёта по данным  WoS CC и исходных данных для показателя Р1(с1) и Р1(с2) программы ПРИОРИТЕТ 2020

(с) Интеграция и Консалтинг 

Программист Е.В. Речкалов, постановка задачи и тестирование М.А. Акоев (Лаборатория Наукометрии, УрФУ)



### Запуск:
```
python __main__.py <template> <...args>
```
где `<template>` - файл шаблона, например `WoS_template_v4.xlsx`

**Обязательные данные:**
- `--name_variants name_variant.txt` - Альтернативные названия университета, пример:
```
Sechenov First Moscow State Medical University
1 MOSCOW MED STATE UNIV IM SECHENOV
1M SECHENOV FIRST MOSCOW STATE MED UNIV
1ST MOSCOW MED ACAD
```

**Небязательные данные:**
- `--ris_ahci RIS_AHCI/`
- `--ris RIS/`
- `--wsd 'Web of Science Documents.csv'`
- `--wsd_q1 'Web of Science DocumentsQ1.csv'`
- `--wsd_q2 'Web of Science DocumentsQ2.csv'`
- `--wsd_q3 'Web of Science DocumentsQ3.csv'`
- `--wsd_q4 'Web of Science DocumentsQ4.csv'`

**Результат:**
- `-o result.xlsx` - Генерируемый excel-файл (по умолчанию - `out.xlsx`)

Зависимости: см. [pyproject.toml](pyproject.toml)
