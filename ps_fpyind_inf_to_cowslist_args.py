# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyind_inf_to_cowslist.py wbN0 sheetN0 colidno0 wbN1 sheetN1 colidno1
# wbN0 : AB_cowshistory.xlsx, sheetN0 : ABFarm, colidno0 : 2 (column number fo idno0), 
# wbN1 : AB_cowslist.xlsx, sheetN1 : cowslist, colidno1 : 2 (column number fo idno1)
import sys
import cowslist

wbN0 = sys.argv[1]
sheetN0 = sys.argv[2]
colidno0 = int( sys.argv[3] )
wbN1 = sys.argv[4]
sheetN1 = sys.argv[5]
colidno1 = int( sys.argv[6] )

cowslist.fpyind_inf_to_cowslist(wbN0, sheetN0, colidno0, wbN1, sheetN1, colidno1)

print( wbN0+ "/"  + sheetN0 + "の個体リストの個体異動情報を検索し、個体情報を" +  wbN1+ "/" + sheetN1  + " に追加入力しました。")
 