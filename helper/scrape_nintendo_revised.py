#! /usr/bin/env python3
'''
Scrape data from the Nintendo Revised Google Sheet
'''

# imports
from io import BytesIO
from openpyxl import load_workbook
from os import makedirs
from os.path import isdir
from urllib.request import urlopen
from warnings import warn

# constants
NINTENDO_REVISED_XLSX_URL = 'https://docs.google.com/spreadsheets/d/1CEABCBrPv1tWf89hSZqUunK0JW-sQo8XpxuvZhdtHQs/export?format=xlsx'
NINTENDO_REVISED_XLSX_SHEETS = {
    '1st Party Nintendo'.upper(),
    'Third PartyIndies'.upper(),
    'Complete Cartridges'.upper(),
}

# main program
if __name__ == "__main__":
    wb = load_workbook(BytesIO(urlopen(NINTENDO_REVISED_XLSX_URL).read()), data_only=True)
    for ws in wb:
        if ws.title.strip().replace('/','').upper() not in NINTENDO_REVISED_XLSX_SHEETS:
            continue
        for row in ws:
            if row[0].value is None or row[6].value is None or row[0].value.strip().upper().startswith('GAME NAME'):
                continue # not a game row
            serial = row[6].value.strip().upper()
            region = serial.split('-')[-1].strip()
            if isdir(serial):
                continue # duplicate
            makedirs(serial)
            f = open('%s/title.txt' % serial, 'w'); f.write('%s\n' % row[0].value.strip()); f.close()
            f = open('%s/region.txt' % serial, 'w'); f.write('%s\n' % region); f.close()
            if row[5].value is not None and row[5].value.strip().upper() != region:
                warn("Mismatch: %s (%s)" % (serial, row[5].value.strip().upper()))
