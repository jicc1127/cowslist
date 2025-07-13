# -*- coding: utf-8 -*-
#Cowslist.xlsx/ABFarmの　2列cowidNoをkeyにして,
#AIData.xlsx/ABFarm 4列 DHINoを検索して
#Cowslist.xlsx/ABFarmに入力します。
#コマンドラインから、引数を渡す
#PS> python ps_fpyinput_eartagno_frm_idno_args.py 
#wb0N, sheet0N, col0, col0_eartag, wb1N, sheet1N, col1, col1_eartag
#PS> python ps_fpyinput_listidno_args.py Cowslist.xlsx ABFarm 1
import sys
import cowslist

wb0N = sys.argv[1]
sheet0N = sys.argv[2]
col0 = int(sys.argv[3])
col0_eartag = int(sys.argv[4])
wb1N = sys.argv[5]
sheet1N = sys.argv[6]
col1 = int(sys.argv[7])
col1_eartag = int(sys.argv[8])

cowslist.fpyinput_eartagno_frm_idno(wb0N, sheet0N, col0, col0_eartag, 
                                            wb1N, sheet1N, col1, col1_eartag)

print("Sheet" + sheet0N + " " + str(col0_eartag) + "列に　DHINo を入力しました。")
