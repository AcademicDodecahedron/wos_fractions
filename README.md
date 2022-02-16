## Скрипт для расчёта исходных данных для показателя Р1(с1) и Р1(с2) программы ПРИОРИТЕТ 2020

(с) Интеграция и Консалтинг 

Программист Е.В. Речкалов, постановка задачи и тестирование М.А. Акоев (Лаборатория Наукометрии, УрФУ)



Пример использования Windows:
```
python wos_fractions_main/ WoS_template_v4.xlsx --ris_ahci RIS_AHCI/ --ris RIS/ --wsd "Web of Science Documents.csv" --wsd_q1 "Web of Science DocumentsQ1.csv" --wsd_q2 "Web of Science DocumentsQ2.csv" --wsd_q3 "Web of Science DocumentsQ3.csv" --wsd_q4 "Web of Science DocumentsQ4.csv" --name_variants name_variant.txt -o out.xlsx
```

 linux:
```
python wos_fractions_main/ WoS_template_v4.xlsx \
       --ris_ahci RIS_AHCI/ \
       --ris RIS/ \
       --wsd 'Web of Science Documents.csv' \
       --wsd_q1 'Web of Science DocumentsQ1.csv' \
       --wsd_q2 'Web of Science DocumentsQ2.csv' \
       --wsd_q3 'Web of Science DocumentsQ3.csv' \
       --wsd_q4 'Web of Science DocumentsQ4.csv' \
       --name_variants name_variant.txt \
       -o out.xlsx
```

Зависимости: `pandas`, `openpyxl`
