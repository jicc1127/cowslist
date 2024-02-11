# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpymod_cowslist_args.py wbN sheetN snewN nco
# wbN : cowshistory.xlsx, sheetN : cowslist, snewN :  cowslistyyyymmdd, ncol : 20 
import sys
import cowslist

wbN = sys.argv[1]
sheetN = sys.argv[2]
snewN = sys.argv[3]
ncol = int(sys.argv[4])


cowslist.fpymod_cowslist( wbN, sheetN, snewN, ncol )

print( wbN+ "/"  + sheetN + "のcowslistの変更点を修正しました。")
