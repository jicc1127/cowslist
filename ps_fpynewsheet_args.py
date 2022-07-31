# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpynewsheet_args.py wbN sheetN scolN r
# ex. wbN: ..\AB_cowslist.xlsx, sheetN: cowslistyyyymmdd, columns, r: 1
import sys
import fmstls

wbN = sys.argv[1]
sheetN = sys.argv[2]
scolN = sys.argv[3]
r = int(sys.argv[4])
fmstls.fpyNewSheet(wbN, sheetN, scolN, r)

print("新しいシート" + sheetN + "を作成しました。")
