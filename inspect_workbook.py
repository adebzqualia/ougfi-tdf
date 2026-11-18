"""Export worksheet headers and KPI formula definitions for mapping review.

This diagnostic never executes formulas or changes a source/history file.
Only the configured Base header row is exported, never contract records.
"""
import argparse
import hashlib
from pathlib import Path
import re
import openpyxl
from pb_automation.common import atomic_json


def inspect(path,base_sheet='Base',header_row=4):
    path=Path(path)
    if header_row<1:raise ValueError('Header row must be positive')
    workbook=openpyxl.load_workbook(path,read_only=True,data_only=False)
    try:
        result={'file':path.name,'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
                'sheet_names':workbook.sheetnames,'base':None,'kpi_formulas':{},
                'scan_limit':'KPI rows 1–200, columns 1–80; Base configured header only'}
        matches=[s for s in workbook.sheetnames if s.casefold()==base_sheet.casefold()]
        if len(matches)==1:
            sh=workbook[matches[0]]
            result['base']={'sheet':sh.title,'header_row':header_row,
                'headers':{c.coordinate:str(c.value) for row in sh.iter_rows(min_row=header_row,max_row=header_row,max_col=80)
                           for c in row if c.value is not None}}
        for sh in workbook:
            if not re.search(r'KPI_?[123]$',sh.title,re.I):continue
            result['kpi_formulas'][sh.title]={c.coordinate:c.value
                for row in sh.iter_rows(max_row=200,max_col=80) for c in row if c.data_type=='f'}
        return result
    finally:workbook.close()


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input',required=True,type=Path)
    parser.add_argument('--output',required=True,type=Path)
    parser.add_argument('--base-sheet',default='Base')
    parser.add_argument('--base-header-row',type=int,default=4)
    args=parser.parse_args()
    if args.input.resolve()==args.output.resolve():parser.error('Output must differ from the input workbook')
    atomic_json(args.output,inspect(args.input,args.base_sheet,args.base_header_row))
    print(f'Header/formula inventory saved: {args.output}')


if __name__=='__main__':main()
