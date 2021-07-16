from argparse import ArgumentParser
from pathlib import Path
from openpyxl import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter
import reader
import c1_parser

def parse_args():
    argp = ArgumentParser()
    argp.add_argument('main_folder', type=Path)
    argp.add_argument('ahci_folder', type=Path)
    return argp.parse_args()
args = parse_args()

# xslx страницы
wb = Workbook()
sel_criteria = wb.active
sel_criteria.title = 'критерийотбора'
sel_criteria.append(['UT', 'критерийотбора'])

publications = wb.create_sheet('публикации')
publications.append(['UT', 'count_au'])

fractions = wb.create_sheet('фракции')
fractions.append(['UT', 'aff', 'org'])

#name_variants = wb.create_sheet('Варианты названий университета')
#name_variants.append(['UT', 'aff', 'org'])


# уникальные UT изи папки AHCI
ut_set = set()
for ahci_file in args.ahci_folder.iterdir():
    print(f"Reading {ahci_file}")
    for record in reader.read(ahci_file):
        ut = record['UT'][0]
        ut_set.add(ut)

# критерии отбора
for ut in ut_set:
    sel_criteria.append([ut, 'Q1 Q2 AHCI'])

# основная папка
for main_file in args.main_folder.iterdir():
    print(f"Reading {main_file}")
    for record in reader.read(main_file):
        ut = record['UT'][0]
        dt = record['DT'][0]
        pt = record['PT'][0]
        publications.append([ut, dt, pt]) # публикации

        af_list = record['AF']
        af_len = len(af_list)
        c1_dict = c1_parser.parse(record['C1'])

        for af in af_list: # каждый автор из AF
            c1_affiliated_unis = c1_dict.get(af, [])
            c1_unis_len = len(c1_affiliated_unis)

            for c1_uni in c1_affiliated_unis: # каждый университет из C1, связанный с автором
                fractions.append([ut, af, af_len, c1_unis_len, c1_uni]) # фракции

def make_table(sheet, name):
    table = Table(displayName=name, ref="A1:" + get_column_letter(sheet.max_column) + str(sheet.max_row))
    sheet.add_table(table)

make_table(sel_criteria, 'критерийотбора')
make_table(publications, 'публикации')
make_table(fractions, 'фракции')
wb.save(filename='out.xlsx')
