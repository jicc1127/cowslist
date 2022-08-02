# -*- coding: utf-8 -*-
import openpyxl
import csv
#import datetime

#fpyopenxl#
"""
fpyopenxl(wbN, sheetN):
    Excelfile wbN.xlsx　sheet sheetN Open 
    v1.00
    2022/1/5
    @author: jicc
    returnがリストでうまく使用できるか不明?2022/02/11
    list[]に出力すれば使用可だが、普通にopenするのとどう違うか疑問
    2022/2/17
"""
def fpyopenxl(wbN, sheetN):
    """
    Excelfile wbN.xlsx　sheet sheetN Open

    Parameters
    ----------
    wbN : str
        ExcelFile Name   ex.MH_CowHistory.xlsx
    sheetN : str
        sheet name

    Returns
    -------
    None.

    """
    
    #import openpyxl
    
    wb = openpyxl.load_workbook(wbN)
    sheet = wb[sheetN]
    return [wb, sheet]

#fpyopencsv_robj#
"""
fpyopencsv_robj:
    csvfile Open for Reader object
    v1.00
    2022/1/5
    @author: jicc
"""
def fpyopencsv_robj(csvN):
    """
    csvfile Open for Reader object

    Parameters
    ----------
    csvN : str
        csvFile Name   ex.MH_??_History.csv

    Returns
    -------
    None.

    """
    #import csv
    
    #filename = csvN.split('.')
    #filename = filename[0]  #拡張子を削除したfilename
    
    filename_file = open(csvN)     #csvfile open
    filename_reader = csv.reader(filename_file)       #get Reader object
    
    
    return filename_reader
    
    
#fpyopencsv_rdata#
"""
fpyopencsv_rdata:
    csvfile Open for Reader data
    v1.00
    2022/1/5
    @author: jicc
"""
def fpyopencsv_rdata(csvN):
    """
    csvfile Open for Reader data

    Parameters
    ----------
    csvN : str
        csvFile Name   ex.MH_??_History.csv

    Returns
    -------
    None.

    """
    #import csv
    
    #filename = csvN.split('.')
    #filename = filename[0]  #拡張子を削除したfilename
    
    filename_file = open(csvN)     #csvfile open
    filename_reader = csv.reader(filename_file)       #get Reader object
    filename_data = list(filename_reader)             #list's list
    
    return filename_data
    
"""
fpyopencsv_w:
    csvfile Open for Writer
    v1.00
    2022/1/7
    @author: jicc
"""
def fpyopencsv_w(csvN):
    """
    csvfile Open for Writer

    Parameters
    ----------
    csvN : str
        csvFile Name   ex.MH_??_History.csv

    Returns
    -------
    None.

    """
    #import csv
    
    output_file = open(csvN, 'w', newline='')       #csvfile open
    output_writer = csv.writer(output_file)       #get Reader object
     
    return output_writer


#fpygetCell_value#
"""
fpygetCell_value: get value from the target Cell
v1.00
2022/2/4

@author: inoue
"""
def fpygetCell_value(sheet, r, col):
    """
    get value from the target Cell on an Excelsheet

    Parameters
    ----------
    sheet : worksheet
        sheetBLV
    r : int
        row
    col : int
        column

    Returns
    -------
    value

    """

    value = sheet.cell(row=r, column=col).value
    return value


#fpyinputCell_value#
"""
fpyinputCell_value: input value to the target Cell
v1.00
2022/2/4

@author: inoue
"""
def fpyinputCell_value(sheet, r, col, vl):
    """
    input value to the target Cell

    Parameters
    ----------
    sheet : worksheet
        sheetBLV
    r : int
        row
    col : int
        column
    vl : type of value
    
    Returns
    -------
    None.

    """

    sheet.cell(row=r, column=col).value = vl    
 
#fpyifNone_inputCell_value#
"""
fpyifNone_inputCell_value:
    if Cellvalue is None,  input value to the Cell
    v1.00
    2022/2/5

    @author: inoue
    
"""
def fpyifNone_inputCell_value(sheet, r, col, vl):
    """
    input value to the target Cell

    Parameters
    ----------
    sheet : worksheet
        sheetBLV
    r : int
        row
    col : int
        column
    vl : type of value
    
    Returns
    -------
    None.

    """
    Cellvalue = sheet.cell(row=r, column=col).value
    if Cellvalue == None:
        sheet.cell(row=r, column=col).value = vl    

#fpyidNo_9to10#
"""
fpyidNo_9to10 : ９～10桁耳標の数値を文字列として、
    9桁の耳標に1桁目に０を加え10桁とする
ｖ1.0
2021/4/29
@author: jicc
"""

#! python3
def fpyidNo_9to10( wbN, sheetN, col ):
    """
    Parameters
    ----------
    wbN : 　str  
        対象となるExcelファイル名　拡張子.xlsxをつけ、　''でくくる。
    sheetN : str
        対象となるsheet
    col : int
        変更する10桁耳標の列

    Returns
    -------
    None.

    """
       
    #import openpyxl
    #import datetime
    
    wb = openpyxl.load_workbook(wbN)
    sheetN = wb[sheetN]   #wb.get_sheet_by_name(sheetN)
    max_row = sheetN.max_row
        
    for row_num in range(2, max_row + 1):     #先頭行をスキップ
        
        idNo = sheetN.cell(row=row_num, column=col).value
        idNo = str(idNo)
        if len(idNo) == 9:
            sheetN.cell(row=row_num, column=col).value = '0' + idNo 
        else:
            sheetN.cell(row=row_num, column=col).value = idNo 
              
    
    wb.save(wbN)
    
    
#fpyNewSheet#
"""
fpyNewSheet : Excelbookに
sheet　'columns'と同じ sheet　'scolN'を作成する。
ｖ1.01
2022/5/3

@author: jicc

"""
def fpyNewSheet(wbN, sheetN, scolN, r):
    """
    Excelbookに sheet 'scolN' r行目の'columns'を1行目に配置した sheet'sheetN'を作成する。
    *sheet 'columns'(列名を記入したシート) を作成しておく
    Parameters
    ----------
    wbN : 　str          
        sheetを作成するワークブック
    sheetN : str　　　　　　シート名:"????" 
        作成するシート
    scolN : str         シート名: "columns"
        参照するシート
	r : int		r行目 作成するcolumn行
    Returns
    -------
    None.

    """
    #import openpyxl
    
    wb = openpyxl.load_workbook(wbN)
    #sheetN = wb[sheetN]
    wb.create_sheet(title=sheetN, index=0)
    sheet = wb[sheetN]
    scol = wb[scolN]
    
    maxcol = scol.max_column #sheet columnの最終列
    
    for i in range(1,maxcol+1):
        sheet.cell(row=r, column=i).value = scol.cell(row=1, column=i).value
    
     
    wb.save(wbN)

"""
fpychgSheetTitle      :change ExcelSheet's title
v1.0
2022/3/30

@author: inoue
"""
def fpychgSheetTitle(wbN, sheetN, sheetN1):
    """
    change the sheet's title

    Parameters
    ----------
    wbN : str
        Excelfile to check double data  '??_CowsHistory.xlsx'
    sheetN : str
        元のシート名  : 'KTFarm'
    sheetN1 : str
        変更名      : 'KTFarmorg' 

    Returns
    -------
    None.

    """
    wbobj = fpyopenxl(wbN, sheetN)
    wb = wbobj[0]
    sheet = wbobj[1]
    sheet.title = sheetN1
    wb.save(wbN)

######################################################################
def fpyfmstlsReference():
    
    print('-----fmstlsReference ---------------------------------------------------------v1.0-------')
    print('**fpyopenxl(wbN, sheetN)')
    print('Excelfile wbN.xlsx　sheet sheetN Open ')
    print('.............................................................................................')
    print('**fpyopencsv_robj(csvN)')
    print('csvfile Open for Reader object')
    print('.............................................................................................')
    print('**fpyopencsv_rdata(csvN)')
    print('csvfile Open for Reader data')
    print('.............................................................................................')
    print('**fpyopencsv_w(csvN)')
    print('csvfile Open for Writer')
    print('.............................................................................................')
    print('**fpygetCell_value(sheet, r, col)')
    print('Excelシート上のセルの値を取得する')
    print('get value from the target Cell on an Excelsheet')
    print('....................................................................................')
    print('**fpyinputCell_value(sheet, r, col, vl)')
    print('Excelシート上のセルに値を上書き入力する')
    print('input value to the target Cell on an Excelsheet')
    print('....................................................................................')
    print('**fpyifNone_inputCell_value(sheet, r, col, vl)')
    print('Excelシート上のセルに値がなければ、入力する')
    print('if Cellvalue is None,  input value to the Cell')
    print('....................................................................................')
    print('**fpyidNo_9to10(wbN, sheetN, col)')
    print('9桁耳標を10桁にし、文字列として再入力する')
    print(' wbN:workbooks_name,  sheetN:worksheets_name, col: columns_no')
    print('....................................................................................')
    print('**fpyNewSheet(wbN, sheetN, scolN, r)')
    print('Excelbookに sheet　\'columns\'r行と同じ sheet　\'scolN\'を作成する')
    print(' wbN:workbooks_name,  sheetN:worksheets_name, scolN: columns_sheets_name')
    print('....................................................................................')
    print('**fpychgSheetTitle(wbN, sheetN, sheetN1)')
    print('change ExcelSheet\'s title')
    print('....................................................................................')
    print('--------------------------------------------------------------------2022/5/4 by jicc---------')
    