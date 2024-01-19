# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpysepclst_outfrmin_args.py wbN sheetN ncol index
# wbN : cowshistory.xlsx, sheetN : ABFarm, ncol : 20, index : 15, 
import sys
import cowslist

wbN = sys.argv[1]
sheetN = sys.argv[2]
ncol = int(sys.argv[3])
index = int(sys.argv[4])


cowslist.fpysepclst_outfrmin( wbN, sheetN, ncol, index )

print( wbN+ "/"  + sheetN + "のcowslistを基準日において 所属牛"+ sheetN+"in" "と 転出牛" + sheetN+"out"  + " に分けました")
