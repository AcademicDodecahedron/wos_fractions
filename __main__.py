from argparse import ArgumentParser
from pathlib import Path
from openpyxl import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter
import reader
import c1_parser
from name_variant import NameVariants, NameVariantsDict

def parse_args():
    argp = ArgumentParser()
    argp.add_argument('main_folder', type=Path)
    argp.add_argument('ahci_folder', type=Path)
    argp.add_argument('name_variants', type=Path)
    argp.add_argument('-o', '--output', default='out.xlsx')
    return argp.parse_args()
args = parse_args()

# xslx страницы
wb = Workbook()
sel_criteria = wb.active
sel_criteria.title = 'критерийотбора'
sel_criteria.append(['UT', 'критерийотбора'])

publications = wb.create_sheet('публикации')
publications.append(['UT', 'Document Type', 'Document Source'])

fractions = wb.create_sheet('фракции')
fractions.append(['UT', 'au', 'count_au', 'count_au_aff', 'aff', 'org'])

name_variants_sheet = wb.create_sheet('Варианты названий университета')
name_variants_sheet.append(['aff', 'org'])

# <-уникальные UT изи папки AHCI
ut_set = set()
for ahci_file in args.ahci_folder.iterdir():
    print(f"Reading {ahci_file}")
    for record in reader.read(ahci_file):
        ut = record['UT'][0]
        ut_set.add(ut)

# ->критерии отбора
for ut in ut_set:
    sel_criteria.append([ut, 'Q1 Q2 AHCI'])

# <-name_variant
name_dict = NameVariantsDict(NameVariants.from_file(args.name_variants))

# <-основная папка
for main_file in args.main_folder.iterdir():
    print(f"Reading {main_file}")
    for record in reader.read(main_file):
        ut = record['UT'][0]
        dt = record['DT'][0]
        pt = record['PT'][0]
        publications.append([ut, dt, pt]) # ->публикации

        af_list = record['AF']
        af_len = len(af_list)
        c1_dict = c1_parser.parse(record['C1'])

        for af in af_list: # каждый автор из AF
            c1_affiliated_unis = c1_dict.get(af, [])
            c1_unis_len = len(c1_affiliated_unis)

            for c1_uni in c1_affiliated_unis: # каждый университет из C1, связанный с автором
                main_uni_name = name_dict.find_and_remember(c1_uni)
                fractions.append([ut, af, af_len, c1_unis_len, c1_uni, main_uni_name]) # ->фракции

#->Варианты названий
for aff, uni in name_dict.get_items():
    name_variants_sheet.append([aff, uni])

def make_table(sheet, name):
    table = Table(displayName=name, ref="A1:" + get_column_letter(sheet.max_column) + str(sheet.max_row))
    sheet.add_table(table)

make_table(sel_criteria, 'критерий_отбора')
make_table(publications, 'публикации')
make_table(fractions, 'фракции')
make_table(name_variants_sheet, 'варианты_названий_университета')
wb.save(filename=args.output)
