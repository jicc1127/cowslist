# -*- coding: utf-8 -*-
import openpyxl
import chghistory
"""
fpymod_cowslist :  modify cowslist data from a new cowslist
    
    v1.0
    2024/2/11
    @author: jicc
    
"""
def fpymod_cowslist( wbN, sheetN, snewN, ncol ):
    """
    modify cowslist data from a new cowslist

    Parameters
    ----------
    wbN : str
        Excelfile to check move-in or move-out data  
        'AB_cowslist.xlsx'　対象のエクセルファイル名
    sheetN : str
        sheet name of an original cowslist(registered cows)
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
    
    #リストのエレメント数 : the number of elements of each list
    lxllists = len(xllists)
    lxllists_new = len(xllists_new)
    
    for i in range(0, lxllists):
        for j in range(0,lxllists_new):
            #original idNo == new cows' idNo
            if xllists[i][1] == xllists_new[j][1]:
                #make a comparison between registered and new cows' eack element
                for k in range(2,ncol):
                    if xllists[i][k] != xllists_new[j][k]:
                        #overwrite a new element
                        xllists[i][k] = xllists_new[j][k]
                    else:
                        continue
            else:
                continue
    
    #overwrite a registered cows' sheet
    chghistory.fpylisttoxls_s_ow(xllists, 1, sheet)
    
    wb.save(wbN)   
                

fpymod_cowslist( '..\MH_cowslist.xlsx', 'cowslist2024', 'cowslist20240407', 20 ) 
           


