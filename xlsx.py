#!/usr/bin/env python3

__author__ = 'Charles Van Goethem'
__license__ = 'unlicense'
__version__ = '0.0.1'
__email__ = 'c-vangoethem@chu-montpellier.fr'
__status__ = 'dev'

import openpyxl
import argh
import tqdm


@argh.arg(
    'files', metavar="fileN.xlsx",
    help="List of files to merge with file 1",
    nargs='+', type=str)
@argh.arg('--output', help="Output file to save.")
def mergeXLSX(files, output="merge.xlsx"):
    "Merge ultiple xlsx files in one. Each sheet dispatch in a file."
    wb_final = openpyxl.Workbook()

    for sheet in wb_final.sheetnames:
        std = wb_final[sheet]
        wb_final.remove(std)

    for f in tqdm.tqdm(files, total=len(files)):
        wb_onload = openpyxl.load_workbook(f)
        for sheet in wb_onload:
            ws = wb_final.create_sheet(sheet.title)
            for row in sheet:
                for cell in row:
                    ws[cell.coordinate].value = cell.value
                    ws[cell.coordinate].style = cell.style

    wb_final.save(output)


if __name__ == '__main__':
    parser = argh.ArghParser()
    parser.add_commands([mergeXLSX])
    parser.dispatch()
