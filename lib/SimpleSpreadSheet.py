"""
SimpleSpreadSheet

Read Excel SpreadSheet to comprehensive dataset

DataSet:

{
    'sheet_names' : ['sheet1','sheet2','sheet3],
    'data' : {
         'sheet1' : [
            ['A','B',...], # Row1
            ['D','E',...], # Row2
            ...
         ]
         'sheet2' : ...
    }
}

"""

from openpyxl import load_workbook
from openpyxl import Workbook

def ReadWorkBookXlsX(filename):
    rtn = {
        'sheet_names' : [],
        'data' : {}
    }
    
    wb = load_workbook(filename)
    for _sheetname in wb.sheetnames:
        sheet = wb[_sheetname]
        d = []
        for row in sheet.rows:
            row_data = []
            for cell in row:    
                if cell.value:
                    v = str(cell.value)
                else:
                    v = ''
                row_data.append(v)
            d.append(row_data)
        rtn['sheet_names'].append(_sheetname)
        rtn['data'][_sheetname] = d
    return rtn

def ReadWorkBook(filename):
    if filename.endswith('.xlsx'):
        return ReadWorkBookXlsX(filename)