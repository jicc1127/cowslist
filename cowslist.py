# -*- coding: utf-8 -*-
"""
Tools for a Farm's cowslist operation
    v1.0
    2022/7/30
    by jicc

"""
import openpyxl
import datetime
import chghistory
import fmstls

#fpyDF_cowslist#############################################################
"""
fpyDF_cowslist:
	Excelfile に 
    Umotionの　飼養牛一覧yyyymmdd から
    cowslistyyyymmdd を作成する
    v1.01
    2022/8/2
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
    fillinDate = fillinDate.date()  
    #date only yyyy-mm-dd v1.01 or + ".strftime('%Y/%m/%d')"
    
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
      birthday = sheetorg.cell(row=row_num, column=6).value
      birthday = birthday.date() 
      #date only yyyy-mm-dd v1.01 or + ".strftime('%Y/%m/%d')"
      sheetN.cell(row=row_num, column=7).value = birthday 
          #sheetorg.cell(row=row_num, column=6).value
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

#fpyind_inf_to_cowlists#####################################################    
"""
fpyind_inf_to_cowslist:
    input individual informations fromcowshistory's data
    to cowslist
    v1.0
    2022/8/2
    @author: jicc
    
    
"""

def fpyind_inf_to_cowslist(wbN0, sheetN0, colidno0, wbN1, sheetN1, colidno1):
    """
    input individual informations from cowshistory's data
    to cowslist    

    Parameters
    ----------
    wbN0 : str
        Excelfile  data of trans information
        ex. "AB_cowhistory.xlsx"
    sheetN0 : str
        sheet name
        ex. "ABFarm"
    colidno0 : int
        column number of 'idno0' (sheetN0 )
    wbN1 : str
        Excelfile list of Farm cow's member
        ex. "AB_cowslist.xlsx"
    sheetN1 : str
        sheet name
        ex. "cowslist"
    colidno1 : int
        column number of 'idno1' (sheetN1 new data)

    Returns
    -------
    None.

    """
    #import chghistory
    #import fmstls
    
    wb0obj = chghistory.fpyopenxl(wbN0, sheetN0)
    #wb0 = wb0obj[0]
    sheet0 = wb0obj[1]
    max_row0 = sheet0.max_row
    
    wb1obj = chghistory.fpyopenxl(wbN1, sheetN1)
    wb1 = wb1obj[0]
    sheet1 = wb1obj[1]
    max_row1 = sheet1.max_row
   
    for row1 in range(2, max_row1+1):
        idno1 = fmstls.fpygetCell_value(sheet1, row1, colidno1) 
        
        for row0 in range(2, max_row0+1):
            idno0 = fmstls.fpygetCell_value(sheet0, row0, colidno0)
            if idno1 == idno0:
                xllists = chghistory.fpyxllist_to_indlist_s( sheet0, 11, idno0 )
                #print('xllixts')
                #rint(xllists)

                #breed
                breed = xllists[0][5]
                if breed == "ホルスタイン種":
                    breed = "H"
                elif breed == "黒毛和種":
                    
                    breed = "W"
                else:
                    breed = "unknown"
            
                fmstls.fpyifNone_inputCell_value(sheet1, row1, 6, breed)
        
                #birthday
                birthday = xllists[0][2]
                fmstls.fpyifNone_inputCell_value(sheet1, row1, 7, birthday)
        
                #sex
                sex = xllists[0][3]
                if sex == "メス":
                    sex = "f"
                elif sex == "オス":
                    sex = "m"
                else:
                    sex = "unknown"
        
                fmstls.fpyifNone_inputCell_value(sheet1, row1, 8, sex)
        
                #damidNo
                damidNo = xllists[0][4]
                fmstls.fpyifNone_inputCell_value(sheet1, row1, 11, damidNo)
                
            else:
                continue
    
    wb1.save(wbN1)

#fpyind_trsinf_to_cowslist##################################################
"""
fpyind_trsinf_to_cowslist:
    input individual transfer informations of cowshistory's data
    to cowslist
    v1.0
    2022/8/5
    '検索年月日' -> fillin_date を追加　#*
    v1.01
    2024/1/5
    @author: jicc
    
    
"""

def fpyind_trsinf_to_cowslist(wbN0, sheetN0, 
                              colidno0, wbN1, sheetN1, colidno1, name):
    """
    input individual transfer informations of cowshistory's data
    to cowslist    

    Parameters
    ----------
    wbN0 : str
        Excelfile  data of transfer information
        ex. "AB_cowhistory.xlsx"
    sheetN0 : str
        sheet name
        ex. "ABFarm"
    colidno0 : int
        column number of 'idno0' (sheetN0 )
    wbN1 : str
        Excelfile list of Farm cow's member
        ex. "AB_cowslist.xlsx"
    sheetN1 : str
        sheet name
        ex. "cowslist"
    colidno1 : int
        column number of 'idno1' (sheetN1 new data)
    name : str
        Farm's name etc. "氏名または名称"

    Returns
    -------
    None.

    """
    #import chghistory
    #import fmstls
    
    wb0obj = chghistory.fpyopenxl(wbN0, sheetN0)
    #wb0 = wb0obj[0]
    sheet0 = wb0obj[1]
    max_row0 = sheet0.max_row
    
    wb1obj = chghistory.fpyopenxl(wbN1, sheetN1)
    wb1 = wb1obj[0]
    sheet1 = wb1obj[1]
    max_row1 = sheet1.max_row
   
    for row1 in range(2, max_row1+1):
        idno1 = fmstls.fpygetCell_value(sheet1, row1, colidno1) 
        
        for row0 in range(2, max_row0+1):
            idno0 = fmstls.fpygetCell_value(sheet0, row0, colidno0)
            if idno1 == idno0:
                xllists = chghistory.fpyxllist_to_indlist_s( sheet0, 12, idno0 )
                xllists.sort(key = lambda x:x[8]) 
                #lists' listを 異動年月日 昇順 でsort lambda関数を利用
                #print('xllixts')
                #print(xllists)
                l = len(xllists)
                #l0 = len(xllists[0])
                for i in range(0, l):
                    if xllists[i][10] == name:  #氏名または名称
                        if xllists[i][7] == "出生" : #異動内容
                            #異動年月日->in_date
                            fmstls.fpyinputCell_value(sheet1, row1, 13, xllists[i][8])
                            #異動内容->in_reason
                            fmstls.fpyinputCell_value(sheet1, row1, 14, xllists[i][7])
                            #氏名または名称->from
                            fmstls.fpyinputCell_value(sheet1, row1, 15, xllists[i][10])
                        elif xllists[i][7] == "転入" : #異動内容
                            #異動年月日->in_date
                            fmstls.fpyinputCell_value(sheet1, row1, 13, xllists[i][8])
                            #異動内容->in_reason
                            fmstls.fpyinputCell_value(sheet1, row1, 14, xllists[i][7])
                            if i > 0:
                                #氏名または名称->from
                                fmstls.fpyinputCell_value(sheet1, row1, 15, xllists[i-1][10])
                            else:
                                continue
                            
                            #out_date -> ''  転入した時点で out information clear
                            fmstls.fpyinputCell_value(sheet1, row1, 16, '')
                            #out_reason -> ''
                            fmstls.fpyinputCell_value(sheet1, row1, 17, '')
                            #to -> ''
                            fmstls.fpyinputCell_value(sheet1, row1, 18, '')
                            
                        elif xllists[i][7] == "転出" : #異動内容 
                            #異動年月日->out_date
                            fmstls.fpyinputCell_value(sheet1, row1, 16, xllists[i][8])
                            #異動内容->out_reason
                            fmstls.fpyinputCell_value(sheet1, row1, 17, xllists[i][7])
                            #氏名または名称->to
                            if i < l-1:
                                fmstls.fpyinputCell_value(sheet1, row1, 18, xllists[i+1][10])
                            else:
                                continue
                        elif xllists[i][7] == "死亡" : #異動内容 
                            #異動年月日->out_date
                            fmstls.fpyinputCell_value(sheet1, row1, 16, xllists[i][8])
                            #異動内容->out_reason
                            fmstls.fpyinputCell_value(sheet1, row1, 17, xllists[i][7])
                            #氏名または名称->to
                            fmstls.fpyinputCell_value(sheet1, row1, 18, xllists[i][10])
                        elif xllists[i][7] == "搬入" or "と畜" or "取引": #異動内容 
                            #異動年月日->out_date
                            fmstls.fpyinputCell_value(sheet1, row1, 16, xllists[i][8])
                            #異動内容->out_reason
                            fmstls.fpyinputCell_value(sheet1, row1, 17, xllists[i][7])
                            #氏名または名称->to
                            fmstls.fpyinputCell_value(sheet1, row1, 18, xllists[i][10])
                    
                    #'検索年月日' -> fillin_date
                    fmstls.fpyinputCell_value(sheet1, row1, 20, xllists[i][11])  #*
                    
    wb1.save(wbN1)

#fpysepclst_outfrmin#############################################################
"""
fpysepclst_outfrmin
    separate move-out cows from move-in in a cowslist
    cowslistのExcelfile: AB_cowslist.xlsx の　sheet　ABFarmの情報を
    基準日における所属牛（転入牛move-in)と転出牛(move-out)の情報に分け、
    2枚のsheet ABFarmin, ABFarmout を作成する
    注) 使用前に２枚のsheet sheetN+'in'と sheetN+'out'を作成しておくこと
        chghistory.fpymkxlsheet(wbN, sheetN, scolN, r)
    v1.0
    2024/1/19
    @author: jicc
    
"""
def fpysepclst_outfrmin( wbN, sheetN, ncol, index ):
    """
    separate move-out cows from move-in in a cowslist
    

    Parameters
    ----------
    wbN : str
        Excelfile to check move-in or move-out data  
        'AB_cowshistory.xlsx'　対象のエクセルファイル名
    sheetN : str
        sheet name to separate move-out cows from move-in
        'ABFarm'　対象のエクセルシート名
    ncol : int
        number of columns sheet ABFarm のリストの列数 : 20
    index : int
        index number of an element(out_date)　
        リスト上の　'out_date'のindex番号

    Returns
    -------
    None.

    """

    #import openpyxl
    #import chghistory
    #import datetime

    wb = openpyxl.load_workbook(wbN)
    sheet = wb[sheetN]
    sheetin = wb[sheetN + 'in']
    sheetout = wb[sheetN + 'out']
    
    xllists = chghistory.fpyxllist_to_list_s(sheet, ncol)
    print('xllists')
    print(xllists[0])
    print(xllists[1])
    print(xllists[2])
    
    farmin = [] #default 所属牛（転入牛move-in)
    farmout = [] #default 転出牛(move-out)
    
    lxllists = len(xllists)
    for i in range(0,lxllists):
        
        out_date = xllists[i][index] #out_date
        if type(out_date) == datetime.datetime:
            farmout.append(xllists[i])
        elif out_date == None:
            farmin.append(xllists[i])
        else:
            xllists[i][18] = "check the value of column 16(out_date)"
            #index 18 : note
            farmin.append(xllists[i])

    #print('farmout')
    #for j in range(0,3):
    #    print(farmout[j])
    
    #print('farmin')
    #for k in range(0,3):
    #    print(farmin[k])
    
    lfarmin = len(farmin)
    for j in range(0,lfarmin):
        
        chghistory.fpylisttoxls_s(farmin[j], 1, sheetin)
    
    lfarmout = len(farmout)
    for k in range(0,lfarmout):
        
        chghistory.fpylisttoxls_s(farmout[k], 1, sheetout)
        
                
    wb.save(wbN)    

'''
    AB_cowslist.xlsx 作成のための　PS用マニュアル
    v1.0 by jicc
    2022/7/31 by jicc

Umotionの　飼養牛一覧yyyymmdd から
    cowslistyyyymmdd を作成する

'''    
def fpycowslistManual():
    
    print('-----cowslist Manual-----------------------------------------------------v1.02------')
    print('1. Umotion 飼養牛一覧　父牛掲載yyyymmdd.csv を')
    print(' AB_cowslist.xlsx/sheet\"cowslistyyyymmddorg\"として移行する')
    print(' ')
    print('2.データ移行用の列名だけのsheet\"cowslistyyyymmdd\"を作成する')
    print('  python ps_fpynewsheet_args.py wbN sheetN scolN row')
    print('wbN : AB_cowslist.xlsx, sheetN : cowslistyyyymmdd, scolN : columns, row : 1')
    print(' ')
    print('3.sheet \"cowlistyyyymmdd\" にデータ入力する')
    print('PS> python ps_fpydf_cowslist_args.py wbN sheetorg sheetN fillinDate')
    print('  wbN : AB_cowslist.xlsx, sheetorg : cowslistyyyymmddorg,')
    print(' sheetN : cowslistyyyymmdd, fillinDate : yyyy/mm/dd')
    print(' ')
    print('4.sheet cowslistyyyymmdd 2列 cowidNoを10桁文字列に統一する')
    print(' PS> python ps_fpyidno_9to10_args.py wbN sheetN col')
    print(' wbN : AB_cowslist.xlsx, sheetN : cowslistyyyymmdd, col : 2')    
    print(' ')
    print('5.input individual informations from cowshistory\'s data to cowslist')
    print(' PS> python ps_fpyind_inf_to_cowslist_args.py wbN0 sheetN0 colidno0 wbN1 sheetN1 colidno1')
    print(' wbN0 : AB_cowshistory.xlsx, sheetN0 : ABFarm, colidno0 : 2 (column number fo idno0), ') 
    print(' wbN1 : AB_cowslist.xlsx, sheetN1 : cowslist, colidno1 : 2 (column number fo idno1)')
    print(' ')
    print('6.input individual transfer informations from cowshistory\'s data to cowslist')
    print(' PS> python ps_fpyind_trsinf_to_cowslist_args.py wbN0 sheetN0 colidno0')
    print(' wbN1 sheetN1 colidno1 name')
    print(' wbN0 : AB_cowshistory.xlsx, sheetN0 : ABFarm, colidno0 : 2 (column number fo idno0), ') 
    print(' wbN1 : AB_cowslist.xlsx, sheetN1 : cowslist, colidno1 : 2 (column number fo idno1)')
    print(' name : 氏名または名称')
    print(' ')
    print('---------------------------------------------------------2022/8/5 by jicc---------')
    

#fpycowslistManual00##########################################################    
'''
    fpycowslistManual00:
    AB_cowslist.xlsx 作成のための　PS用マニュアル
    v2.0 by jicc
    2024/1/14 by jicc

    AB_cowshistory.xlsx の個体情報、異動情報から　
    cowslistyyyymmdd を作成する

'''    
def fpycowslistManual00():

    print('-----cowslist Manual00---------------------------------------------------v2.0------')    
    print('1.input individual informations from cowshistory\'s data to cowslist')
    print(' PS> python ps_fpyind_inf_to_cowslist_args.py wbN0 sheetN0 colidno0 wbN1 sheetN1 colidno1')
    print(' wbN0 : AB_cowshistory.xlsx, sheetN0 : ABFarm, colidno0 : 2 (column number fo idno0), ') 
    print(' wbN1 : AB_cowslist.xlsx, sheetN1 : cowslist, colidno1 : 2 (column number fo idno1)')
    print(' ')
    print('2.input individual transfer informations from cowshistory\'s data to cowslist')
    print(' PS> python ps_fpyind_trsinf_to_cowslist_args.py wbN0 sheetN0 colidno0')
    print(' wbN1 sheetN1 colidno1 name')
    print(' wbN0 : AB_cowshistory.xlsx, sheetN0 : ABFarm, colidno0 : 2 (column number fo idno0), ') 
    print(' wbN1 : AB_cowslist.xlsx, sheetN1 : cowslist, colidno1 : 2 (column number fo idno1)')
    print(' name : 氏名または名称')
    print(' ')
    print('3.make an ExcelSheet if it dose not exist')
    print(' PS> python ps_fpymkxlsheet_args.py wbN sheetN scolN r')
    print(' wbN: ..\AB_cowslist.xlsx, sheetN: ABFarmout, scolN: columns, r: 1')
    print(' 2枚のsheet ABFarmin, ABFarmout がなければ、作成する ') 
    print(' ')
    print('4.separate move-out cows from move-in in a cowslist')
    print(' PS> python ps_fpysepclst_outfrmin_args.py wbN sheetN ncol index')
    print(' wbN: ..\AB_cowslist.xlsx, sheetN: ABFarmout, ncol: 20, r: index : 15')
    print(' 基準日における所属牛（転入牛move-in)と転出牛(move-out)の情報に分け、 ') 
    print(' 2枚のsheet ABFarmin, ABFarmout を作成する ')
    print('---------------------------------------------------------2024/1/19 by jicc---------')