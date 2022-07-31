# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyidno_9to10_args.py wbN cowslistyyyymmdd 2(B列)
#wbN: AB_cowslist.xlsx
import sys
import fmstls

wbN = sys.argv[1]
sheetN = sys.argv[2]
col = int(sys.argv[3])
print(sys.argv[2])
print(sys.argv[3])
print(int(sys.argv[3]))
print(col)

fmstls.fpyidNo_9to10( wbN, sheetN, col )
print(" 個体識別番号を10桁文字列に統一しました。")
 
