from lxml import html
import pandas as pd
from openpyxl import load_workbook
from openpyxl import worksheet
from pathlib import Path



def classifyData(type):

    file_source = Path('excel/FIREANT_dscp_' + type + '.xlsx')
    file_result = Path('excel/clsf_' + type + '.xlsx')
    if not file_result.is_file():
        writer_result = pd.ExcelWriter('excel/clsf_' + type + '.xlsx', engine='openpyxl')
    if file_source.is_file():
        book_resource = load_workbook('excel/FIREANT_dscp_' + type + '.xlsx')
        book_result = load_workbook('excel/clsf_' + type + '.xlsx')
        for ws in book_resource.worksheets:
            max_row =  ws.max_row
            max_col = ws.max_column
            for i in range(1,max_row):
                if type == 'Nam':
                    name_sheet_active = ws.cell(i,3)
                    ws.delete_cols()
                else:
                    name_sheet_active = ws.cell(i,3)+'_'+ ws.cell(i,4)
                if name_sheet_active not in book_result.sheetnames:                    
                    book_result.create_sheet(name_sheet_active)
                    book_result.sheetnames(name_sheet_active)

                ws_result = book_result[name_sheet_active]






        #writer_result = pd.ExcelWriter('excel/clsf_' + type + '.xlsx', engine='openpyxl')
        writer_result.book = book_result
        writer_result.sheets = dict((ws.title, ws) for ws in book_resource.worksheets)
    else:
        print('du lieu dau!!!!')

    my_file = Path('excel/FIREANT_' + nhomcp + '_Nam.xlsx')
    if my_file.is_file():
        book_quy = load_workbook('excel/FIREANT_' + nhomcp + '_Quy.xlsx')
        writer_quy = pd.ExcelWriter('excel/FIREANT_' + nhomcp + '_Quy.xlsx', engine='openpyxl')
        writer_quy.book = book_quy
        writer_quy.sheets = dict((ws.title, ws) for ws in book_quy.worksheets)
    else:
        writer_quy = pd.ExcelWriter('excel/FIREANT_' + nhomcp + '_Quy.xlsx', engine='openpyxl')

    for i in ds:
        print('Lay du lieu ',i,' theo nam')
        data = Get_company(i, 'Y', 2017, 2018)
        data.to_excel(writer_nam, sheet_name=i, encoding='utf-8-sig')
        print('Lay du lieu ',i,' theo quy')
        data = Get_company(i, 'Q', 2017, 2018)
        data.to_excel(writer_quy, sheet_name=i, encoding='utf-8-sig')
    writer_nam.save()
    writer_quy.save()
    print('xong')

run()