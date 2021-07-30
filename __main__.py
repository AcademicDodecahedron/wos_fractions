# (c) E-kvadrat Consulting & Media, 2021

from argparse import ArgumentParser
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter
from typing import NamedTuple, Any, Optional
from collections import defaultdict
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
    argp.add_argument('--name_variants', required=True, type=Path, help='name_variant.txt file')
    argp.add_argument('-o', '--output', default='out.xlsx', help='Output .xlsx file')
    return argp.parse_args()
args = parse_args()

wb = openpyxl.load_workbook(args.template)
sel_criteria = wb['критерийотбора']
publications = wb['публикации']
fractions = wb['фракции']
name_variants_sheet = wb['Варианты названий университета']

# <-name_variant
name_dict = NameVariantsDict(NameVariants.from_file(args.name_variants))
main_name = name_dict.get_main_name()

##
FORMULA_PUBLICATIONS_CRITERIA = '=IFERROR(VLOOKUP(публикации[UT],критерий_отбора[],2,0),"")'
FORMULA_PUBLICATIONS_FRACTION = f'=IFERROR(SUMIFS(фракции[фракции],фракции[org],"={main_name}",фракции[UT],"="&публикации[[#This Row],[UT]]),"")'
FORMULA_FRACTION_FRACTION = '=1/фракции[count_au]/фракции[count_au_aff]'
##
# в каких файлах есть этот UT
class YearRow(NamedTuple):
    year: Any

class WsdRow(NamedTuple):
    year: Any
    document_type: Any

class UtSource:
    def __init__(self):
        self.ris: Optional[YearRow] = None
        self.ris_ahci: Optional[YearRow] = None
        self.wsd: Optional[WsdRow] = None
        self.wsd_q1: Optional[YearRow] = None
        self.wsd_q2: Optional[YearRow] = None
        self.wsd_q3: Optional[YearRow] = None
        self.wsd_q4: Optional[YearRow] = None

    def get_priority_year(self):
        if self.ris_ahci is not None and self.ris_ahci.year is not None:
            return self.ris_ahci.year
        elif self.wsd_q1 is not None and self.wsd_q1.year is not None:
            return self.wsd_q1.year
        elif self.wsd_q2 is not None and self.wsd_q2.year is not None:
            return self.wsd_q2.year
        elif self.wsd_q3 is not None and self.wsd_q3.year is not None:
            return self.wsd_q3.year
        elif self.wsd_q4 is not None and self.wsd_q4.year is not None:
            return self.wsd_q4.year
        elif self.wsd is not None and self.wsd.year is not None:
            return self.wsd.year
        elif self.ris is not None and self.ris.year is not None:
            return self.ris.year

    def get_document_type_wsd(self):
        if self.wsd is not None:
            return self.wsd.document_type

    def get_quartile(self):
        if self.ris_ahci is not None:
            return 'AHCI'
        elif self.wsd_q1 is not None:
            return 'Q1'
        elif self.wsd_q2 is not None:
            return 'Q2'
        elif self.wsd_q3 is not None:
            return 'Q3'
        elif self.wsd_q4 is not None:
            return 'Q4'
        else:
            return 'n/a'

ut_sources = defaultdict(UtSource)
##

def try_get_first_line(dictionary: dict, key: str):
    lines = dictionary.get(key, None)
    if lines is not None:
        return lines[0]

# <-RIS_AHCI
if args.ris_ahci is not None:
    for ahci_file in args.ris_ahci.iterdir():
        print('Reading ', ahci_file)
        for record in reader.read(ahci_file):
            ut = record['UT'][0]
            year = try_get_first_line(record, 'PY')
            ut_sources[ut].ris_ahci = YearRow(year) #->источники

def each_ut_and_year_from_csv(path):
    if path is not None:
        print('Reading ', path)
        for row in csv_utils.read_csv_body(path):
            ut = row['Accession Number']
            year = row['Publication Date']
            yield ut, year
    else:
        return []

# <-Web Science Of DocumentsQ*.csv
for ut, year in each_ut_and_year_from_csv(args.wsd_q1):
    ut_sources[ut].wsd_q1 = YearRow(year) #->источники
for ut, year in each_ut_and_year_from_csv(args.wsd_q2):
    ut_sources[ut].wsd_q2 = YearRow(year) #->источники
for ut, year in each_ut_and_year_from_csv(args.wsd_q3):
    ut_sources[ut].wsd_q3 = YearRow(year) #->источники
for ut, year in each_ut_and_year_from_csv(args.wsd_q4):
    ut_sources[ut].wsd_q4 = YearRow(year) #->источники

wsd_ut_set = set()
#<-Web Science Of Document.csv
if args.wsd is not None:
    print('Reading ', args.wsd)
    for i, row in enumerate(csv_utils.read_csv_body(args.wsd)):
        ut = row['Accession Number']
        document_type = row['Document Type']
        year = row['Publication Date']

        wsd_ut_set.add(ut)
        ut_sources[ut].wsd = WsdRow(year, document_type)

def append_fractions(ut, au, count_au, count_au_aff, aff):
    org = 'NoN'
    name_aff_variant = name_dict.find_matching_pattern(aff)
    if name_aff_variant is not None:
        org = main_name
    fractions.append([
        ut,
        au,
        count_au,
        count_au_aff,
        aff,
        org,
        name_aff_variant,
        FORMULA_FRACTION_FRACTION
    ]) # ->фракции

# <-RIS
if args.ris is not None:
    for ris_file in args.ris.iterdir():
        print('Reading ', ris_file)
        for record in reader.read(ris_file):
            ut = record['UT'][0]
            year = try_get_first_line(record, 'PY')
            ut_source = ut_sources[ut]
            ut_source.ris = YearRow(year)
            wsd_ut_set.discard(ut)

            document_type = record['DT'][0]
            document_source = record['PT'][0]

            publications.append([
                ut,
                ut_source.get_priority_year(),
                document_type,
                ut_source.get_document_type_wsd(),
                document_source,
                FORMULA_PUBLICATIONS_CRITERIA,
                FORMULA_PUBLICATIONS_FRACTION
            ]) #->публикации

            au_list = record['AF']
            count_au = len(au_list)
            c1_data = c1_parser.parse(record['C1'])

            if type(c1_data) == dict:
                # В C1 данный вида [Автор] Университет
                for au in au_list: # каждый автор из AF
                    c1_affiliated_unis = c1_data.get(au, []) #type:ignore
                    count_au_aff = len(c1_affiliated_unis)

                    for aff in c1_affiliated_unis: # каждый университет из C1, связанный с автором
                        append_fractions(ut, au, count_au, count_au_aff, aff)

            elif len(au_list) == 1:
                # В C1 одна строка без указания автора, значит автор в AF должен быть один
                append_fractions(ut, au_list[0], count_au, 1, c1_data)
            else:
                print('C1 is a string:', c1_data, 'expected exactly 1 author, got', au_list)
                exit(1)

for ut in wsd_ut_set:
    ut_source = ut_sources[ut]

    publications.append([
        ut,
        ut_source.get_priority_year(),
        None,
        ut_source.get_document_type_wsd(),
        None,
        FORMULA_PUBLICATIONS_CRITERIA,
        FORMULA_PUBLICATIONS_FRACTION
    ]) #->публикации

#<-словарь имен
for aff, _ in name_dict.get_items():
    name_variants_sheet.append([aff, main_name]) #->Варианты названий университета

def get_criteria_name(quartile: str):
    if quartile == 'Q1' or quartile == 'Q2' or quartile == 'AHCI':
        return 'Q1 Q2 AHCI'
    else:
        return 'Q3 Q4 n/a'

#<-источники ut
for ut, sources in ut_sources.items():
    quartile = sources.get_quartile()
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
