# -*- coding: utf-8 -*-
"""
Tools for a Farm's cowslist operation
    v1.0
    2022/7/30
    by jicc

"""
import openpyxl
import datetime

"""
fpyDF_cowslist:
	Excelfile に 
    Umotionの　飼養牛一覧yyyymmdd から
    cowslistyyyymmdd を作成する
    v1.0
    2022/7/30
    @author:jicc

"""
#! python3
def fpyDF_cowslist( wbN, sheetorg, sheetN, fillinDate ):
    """
    Excelfile に 
    Umotionの　飼養牛一覧yyyymmdd から
    cowslistyyyymmdd を作成する
    
    Parameters
    ----------
    wbN : ワークブック名
        "AB_cowslist.xlsx"
    sheetorg : データ参照シート
        "cowslistyyyymmddorg"
    sheetN : 経産牛データシート
        "cowslistyyyymmdd"
    fillinDate : 作成日
        "yyyy/mm/dd"
    
    Returns
    -------
    None.

    """
    
    #import openpyxl
    #import datetime

    wb = openpyxl.load_workbook(wbN)
    sheetorg = wb[sheetorg]

    sheetN = wb[sheetN]
    max_row = sheetorg.max_row
    fillinDate = datetime.datetime.strptime(fillinDate, '%Y/%m/%d')
    
    for row_num in range(2, max_row + 1):     #先頭行をスキップ
    
      #LineNo      
      sheetN.cell(row=row_num, column=1).value = row_num - 1
      
      #cowidNo
      sheetN.cell(row=row_num, column=2).value = sheetorg.cell(row=row_num, 
                                                            column=2).value
                                                  #個体識別番号
      #DHITNo
      sheetN.cell(row=row_num, column=4).value = sheetorg.cell(row=row_num, 
                                                            column=1).value
                                                  #牛番号
      #birthday
      sheetN.cell(row=row_num, column=7).value = sheetorg.cell(row=row_num, 
                                                            column=6).value
                                                  #出生日
      #sire_code     
      sheetN.cell(row=row_num, column=9).value = sheetorg.cell(row=row_num, 
                                                            column=14).value
                                                  #父牛の略号
      #sire_name
      sheetN.cell(row=row_num, column=10).value = sheetorg.cell(row=row_num, 
                                                            column=13).value
                                                  #父牛の名前
      #fillindate
      sheetN.cell(row=row_num, column=20).value = fillinDate
      
                                                
    
    wb.save(wbN)

'''
    AB_cowslist.xlsx 作成のための　PS用マニュアル
    v1.0 by jicc
    2022/7/31 by jicc

Umotionの　飼養牛一覧yyyymmdd から
    cowslistyyyymmdd を作成する

'''    
def fpycowslistManual():
    
    print('-----cowslist Manual-----------------------------------------------------v1.0-------')
    print('1. Umotion 飼養牛一覧　父牛掲載yyyymmdd.csv を')
    print(' AB_cowslist.xlsx/sheet\"cowslistyyyymmddorg\"として移行する')
    print(' ')
    print('2.データ移行用の列名だけのsheet\"cowslistyyyymmdd\"を作成する')
    print('  python ps_fpynewsheet_args.py wbN sheetN scolN row')
    print('wbN : AB_cowslist.xlsx, sheetN : cowslistyyyymmdd, scolN : columns, row : 1')
    print(' ')
    print('3.sheet \"cowlistyyyymmdd\" にデータ入力する')
    print('PS> python ps_fpydf_cowslist_args.py wbN sheetorg sheetN fillinDate')
    print('  wbN : AB_cowslist.xlsx, sheetorg : cowslistyyyymmddorg,'
    print(' sheetN : cowslistyyyymmdd, fillinDate : yyyy/mm/dd')
    print(' ')
    print('4.sheet cowslistyyyymmdd 2列 cowidNoを10桁文字列に統一する')
    print(' PS> python ps_fpyidno_9to10_args.py wbN sheetN col')
    print(' wbN : AB_cowslist.xlsx, sheetN : cowslistyyyymmdd, col : 2')    
    print(' ')
    print('---------------------------------------------------------2022/7/31 by jicc---------')
    