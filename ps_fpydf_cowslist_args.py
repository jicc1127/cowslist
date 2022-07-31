# -*- coding: utf-8 -*-
#sheet"yyyymmddCow00"にデータ入力します。
#コマンドラインから、引数を渡す
#　PS> python ps_fpydf_cowslist_args.py wbN cowslistyyyymmddorg cowslistyyyymmdd yyyy/mm/dd
import sys
import cowslist

wbN = sys.argv[1]
sheetorg = sys.argv[2]
sheetN = sys.argv[3]
fillinDate = sys.argv[4]

cowslist.fpyDF_cowslist(wbN, sheetorg, sheetN, fillinDate)

print("Sheet" + sheetN + "に　データを入力しました。")
