# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyreg_newcows_args.py wbN sheetN snewN nco
# wbN : cowshistory.xlsx, sheetN : cowslist, snewN :  cowslistyyyymmdd, ncol : 20 
import sys
import cowslist

wbN = sys.argv[1]
sheetN = sys.argv[2]
snewN = sys.argv[3]
ncol = int(sys.argv[4])


cowslist.fpyreg_newcows( wbN, sheetN, snewN, ncol )

print( wbN+ "/"  + sheetN + "のcowslistに新規牛を追加登録しました。")
