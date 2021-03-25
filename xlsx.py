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
def merge_files(files, output="merge.xlsx"):
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


@argh.arg(
    'file', metavar="file.xlsx", type=str,
    help="File wich header will be removed (default: active sheet/1st line).")
@argh.arg(
    '--sheets', metavar="name",
    help="List of sheet names wich header will be removed.",
    nargs='*', type=str)
@argh.arg('--all-sheets', help="Remove header to all sheets.")
@argh.arg(
    '--lines',
    help="Number of row of the header.", nargs='+', type=int)
@argh.arg('--output', help="Output file (default same file).")
def remove_lines(file, sheets=None, all_sheets=False, lines=[1], output=None):
    "Remove line(s) from sheet(s)."
    workbook = openpyxl.load_workbook(file)

    list_of_sheets = [workbook.active.title]
    if sheets is not None:
        list_of_sheets = sheets
    if all_sheets:
        list_of_sheets = workbook.sheetnames

    for sheet in list_of_sheets:
        for line in lines:
            workbook[sheet].delete_rows(line)
    out_file = str()
    if output is None:
        out_file = file
    else:
        out_file = output

    workbook.save(out_file)


if __name__ == '__main__':
    parser = argh.ArghParser()
    parser.add_commands([merge_files, remove_lines])
    parser.dispatch()
