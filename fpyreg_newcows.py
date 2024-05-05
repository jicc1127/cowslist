# -*- coding: utf-8 -*-
import openpyxl
import chghistory
"""
fpyreg_newcows : register new cows from a new cowslist
    
    v1.0
    2024/2/11
    @author: jicc
    
"""
def fpyreg_newcows( wbN, sheetN, snewN, ncol ):
    """
    register new cows from a new cowslist

    Parameters
    ----------
    wbN : str
        Excelfile to check move-in or move-out data  
        'AB_cowslist.xlsx'　対象のエクセルファイル名
    sheetN : str
        sheet name of an original cowslist
        'cowslist', 'cowslist2024' etc 追加する対象のエクセルシート名
    snewN : str
        sheet name of a new cowslist
        'cowslistyyyymmdd'
    ncol : int
        number of columns sheet cowslist のリストの列数 : 20

    Returns
    -------
    None

    """
    #import openpyxl
    #import chghistory
    
    wb = openpyxl.load_workbook(wbN)
    sheet = wb[sheetN]
    snew = wb[snewN]
    
    #sheet と　snew の Excel sheetを　それぞれ listにする
    xllists = chghistory.fpyxllist_to_list_s(sheet, ncol)
    xllists_new = chghistory.fpyxllist_to_list_s(snew, ncol)
    
    #リストのエレメント数 : the number of each list
    lxllists = len(xllists)
    lxllists_new = len(xllists_new)
    
    #振り分け用のlistの初期化　: 
    xllists0 = []  #new cows list default
    for h in range(0,lxllists_new):
        xllists0.append(xllists_new[h])
    #xllists0 = xllists_new とすると、
    #xllists0, xllists_new ともにremoveされ　index errorが起こる 2024/2/9
    
    xllists1 = []   #registered cows list default
    
    #xllists01 = [xllists0, xllists1]
    
    #登録牛と新規牛を分ける
    #separate new cows from registered cows
    for i in range(0, lxllists):
        for j in range(0,lxllists_new):
            if xllists[i][1] == xllists_new[j][1]: #registered
                #the idNo of xllixts[i]  equal th idNo of xllists_new[j]
                xllists1.append(xllists_new[j]) #未使用
                xllists0.remove(xllists_new[j])
                break
            else:
                continue
    
    print('xllists0')
    l0 = len(xllists0)
    print(l0)
    print(xllists0)
    print('xllists1')
    l1 = len(xllists1)
    print(l1)
    xllists1[1]
    
    #新規牛のcowlist_id を修正する : modify cowlist_id of xllists0[k][0]
    l = lxllists+1
    for k in range(0,l0):
        xllists0[k][0] = l
        l = l + 1
    
    #新規牛をExcelsheet cowslist追加 : add new cows to a cowslist
    chghistory.fpylisttoxls_s_(xllists0, 1, sheet)
    
    wb.save(wbN)   
                

fpyreg_newcows( '..\MH_cowslist.xlsx', 'cowslist2024', 'cowslist20240407', 20 ) 
           


