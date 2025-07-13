# -*- coding: utf-8 -*-
#Cowslist.xlsx/
#sheet"ABFarm"の1列　cowlist_idを入力します。
#コマンドラインから、引数を渡す
#   PS> python ps_fpyinput_listidno_args.py WbN sheetN col
#　PS> python ps_fpyinput_listidno_args.py Cowslist.xlsx ABFarm 1
import sys
import fmstls

wbN = sys.argv[1]
sheetN = sys.argv[2]
col = int(sys.argv[3])

fmstls.fpynumber_rows( wbN, sheetN, col )

print("Sheet" + sheetN + " " + str(col) + "列に　cowlist_id を入力しました。")
