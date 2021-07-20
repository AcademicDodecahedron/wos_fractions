from argparse import ArgumentParser
from pathlib import Path
import openpyxl
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter
from collections import defaultdict
import reader
import c1_parser
from name_variant import NameVariants, NameVariantsDict
import csv_utils

def parse_args():
    argp = ArgumentParser()
    argp.add_argument('template', type=Path, help='.xlsx template file')
    argp.add_argument('--ris', type=Path, help='RIS folder')
    argp.add_argument('--ris_ahci', type=Path, help='RIS_AHCI folder')
    argp.add_argument('--wsd', type=Path, help='Web of Science Documents.csv')
    argp.add_argument('--wsd_q1', type=Path, help='Web of Science DocumentsQ1.csv')
    argp.add_argument('--wsd_q2', type=Path, help='Web of Science DocumentsQ2.csv')
    argp.add_argument('--wsd_q3', type=Path, help='Web of Science DocumentsQ3.csv')
    argp.add_argument('--wsd_q4', type=Path, help='Web of Science DocumentsQ4.csv')
    argp.add_argument('--name_variants', type=Path, help='name_variant.txt file')
    argp.add_argument('-o', '--output', default='out.xlsx', help='Output .xlsx file')
    return argp.parse_args()
args = parse_args()

wb = openpyxl.load_workbook(args.template)
sel_criteria = wb['критерийотбора']
publications = wb['публикации']
fractions = wb['фракции']
name_variants_sheet = wb['Варианты названий университета']

# <-уникальные UT изи папки AHCI
class UtSource:
    def __init__(self):
        self.ris = False
        self.ris_ahci = False
        self.wsd = False
        self.wsd_q1 = False
        self.wsd_q2 = False
        self.wsd_q3 = False
        self.wsd_q4 = False

ut_sources = defaultdict(UtSource)
if args.ris_ahci is not None:
    for ahci_file in args.ris_ahci.iterdir():
        print('Reading ', ahci_file)
        for record in reader.read(ahci_file):
            ut = record['UT'][0]
            ut_sources[ut].ris_ahci = True

if args.ris is not None:
    for ris_file in args.ris.iterdir():
        print('Reading ', ris_file)
        for record in reader.read(ris_file):
            ut = record['UT'][0]
            ut_sources[ut].ris = True

def read_csv_ut_if_some(path):
    if path is not None:
        print('Reading ', path)
        for row in csv_utils.read_csv_body(path):
            ut = row['Accession Number']
            yield ut
    else:
        return []

for ut in read_csv_ut_if_some(args.wsd):
    ut_sources[ut].wsd = True
for ut in read_csv_ut_if_some(args.wsd_q1):
    ut_sources[ut].wsd_q1 = True
for ut in read_csv_ut_if_some(args.wsd_q2):
    ut_sources[ut].wsd_q2 = True
for ut in read_csv_ut_if_some(args.wsd_q3):
    ut_sources[ut].wsd_q3 = True
for ut in read_csv_ut_if_some(args.wsd_q4):
    ut_sources[ut].wsd_q4 = True

def get_quartile_name(sources: UtSource):
    if sources.wsd_q1:
        return 'Q1'
    elif sources.wsd_q2:
        return 'Q2'
    elif sources.wsd_q3:
        return 'Q3'
    elif sources.wsd_q4:
        return 'Q4'
    elif sources.ris_ahci:
        return 'AHCI'
    else:
        return 'n/a'

def get_criteria_name(quartile: str):
    if quartile == 'Q1' or quartile == 'Q2' or quartile == 'AHCI':
        return 'Q1 Q2 AHCI'
    else:
        return 'Q3 Q4 n/a'


# ->критерии отбора
for ut, sources in ut_sources.items():
    quartile = get_quartile_name(sources)
    criteria_name = get_criteria_name(quartile)

    sel_criteria.append([
        ut,
        criteria_name,
        None,
        quartile
    ])

def update_table(sheet, name):
    table = sheet.tables[name]
    table.ref = "A1:" + get_column_letter(sheet.max_column) + str(sheet.max_row)

update_table(sel_criteria, 'критерий_отбора')
wb.save(args.output)
