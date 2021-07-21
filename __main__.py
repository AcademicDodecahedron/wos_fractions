from argparse import ArgumentParser
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter
from collections import defaultdict
import pandas as pd
import reader
import c1_parser
from name_variant import NameVariants, NameVariantsDict
import csv_utils

##
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

##
# в каких файлах есть этот UT
class UtSource:
    def __init__(self):
        self.ris = False
        self.ris_ahci = False
        self.wsd_q1 = False
        self.wsd_q2 = False
        self.wsd_q3 = False
        self.wsd_q4 = False

ut_sources = defaultdict(UtSource)
##

# <-RIS_AHCI
if args.ris_ahci is not None:
    for ahci_file in args.ris_ahci.iterdir():
        print('Reading ', ahci_file)
        for record in reader.read(ahci_file):
            ut = record['UT'][0]
            ut_sources[ut].ris_ahci = True #->источники

def each_ut_from_csv(path):
    if path is not None:
        print('Reading ', path)
        for row in csv_utils.read_csv_body(path):
            ut = row['Accession Number']
            yield ut
    else:
        return []

# <-Web Science Of DocumentsQ*.csv
for ut in each_ut_from_csv(args.wsd_q1):
    ut_sources[ut].wsd_q1 = True #->источники
for ut in each_ut_from_csv(args.wsd_q2):
    ut_sources[ut].wsd_q2 = True #->источники
for ut in each_ut_from_csv(args.wsd_q3):
    ut_sources[ut].wsd_q3 = True #->источники
for ut in each_ut_from_csv(args.wsd_q4):
    ut_sources[ut].wsd_q4 = True #->источники

# <-name_variant
name_dict = None
if args.name_variants is not None:
    name_dict = NameVariantsDict(NameVariants.from_file(args.name_variants))

ris_df = pd.DataFrame(columns=['UT', 'Type', 'Source', 'Year'])

# <-RIS
if args.ris is not None:
    row = 0
    for ris_file in args.ris.iterdir():
        print('Reading ', ris_file)
        for record in reader.read(ris_file):
            ut = record['UT'][0]
            ut_sources[ut].ris = True

            document_type = record['DT'][0]
            document_source = record['PT'][0]
            year = None
            if 'PY' in record:
                year = record['PY'][0]

            ris_df.loc[row] = [ut, document_type, document_source, year] #-> Сохранить в DataFrame
            row += 1

            au_list = record['AF']
            count_au = len(au_list)
            c1_dict = c1_parser.parse(record['C1'])

            for au in au_list: # каждый автор из AF
                c1_affiliated_unis = c1_dict.get(au, [])
                count_au_aff = len(c1_affiliated_unis)

                for aff in c1_affiliated_unis: # каждый университет из C1, связанный с автором
                    org = None
                    if name_dict is not None:
                        org = name_dict.find_and_remember(aff)

                    fractions.append([
                        ut,
                        au,
                        count_au,
                        count_au_aff,
                        aff,
                        org,
                        '=1/фракции[count_au]/фракции[count_au_aff]'
                    ]) # ->фракции

#<-словарь имен
if name_dict is not None:
    for aff, uni in name_dict.get_items():
        name_variants_sheet.append([aff, uni]) #->Варианты названий университета

wsd_df = pd.DataFrame(columns=['UT', 'Type', 'Year'])

#<-Web Science Of Document.csv
if args.wsd is not None:
    print('Reading ', args.wsd)
    for i, row in enumerate(csv_utils.read_csv_body(args.wsd)):
        ut = row['Accession Number']
        document_type = row['Document Type']
        year = row['Publication Date']
        wsd_df.loc[i] = [ut, document_type, year] #-> Сохранить в DataFrame

# Объединить RIS и InCites
publications_df = ris_df.merge(wsd_df, how='outer', on='UT', suffixes=('_ris', '_wsd'))
publications_df['Year_wsd'] = publications_df['Year_wsd'] \
    .combine_first(publications_df['Year_ris']) #type:ignore

#<-DataFrame
for _, row in publications_df.iterrows():
    publications.append([
        row['UT'],
        row['Year_wsd'],
        row['Type_ris'],
        row['Type_wsd'],
        row['Source'],
        '=IFERROR(VLOOKUP(публикации[UT],критерий_отбора[],2,0),"")',
        '=VLOOKUP(публикации[UT],фракции_сводная[],2,0)'
    ]) #->публикации

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

#<-источники ut
for ut, sources in ut_sources.items():
    quartile = get_quartile_name(sources)
    criteria_name = get_criteria_name(quartile)

    sel_criteria.append([
        ut,
        criteria_name,
        None,
        quartile
    ]) #-> критерийотбора

def update_table(sheet, name):
    table = sheet.tables[name]
    table.ref = "A1:" + get_column_letter(sheet.max_column) + str(sheet.max_row)

update_table(sel_criteria, 'критерий_отбора')
update_table(fractions, 'фракции')
update_table(publications, 'публикации')
update_table(name_variants_sheet, 'варианты_названий_университета')
wb.save(args.output)
