# -*- coding: utf-8 -*-
#コマンドラインから、引数を渡す
#　PS> python ps_fpyext_cwslst_at_base_date_args.py wbN0 sN0 coln0 ncol0 index name
# wbN1 sN1 coln1 ncol1 bdate
# wbN0 : AB_cowshistory.xlsx, sN0 : ABFarm, coln0 : 2, ncol0 : 12, index : 10, name : 'Farm name', 
# wbN1 : AB_cowslist.xlsx, sN1 : cowslist2024, coln1 : 2, ncol1 : 20, bdate : 2024/1/14
import sys
import cowslist

wbN0 = sys.argv[1]
sN0 = sys.argv[2]
coln0 = int(sys.argv[3])
ncol0 = int(sys.argv[4])
index = int(sys.argv[5])
name = sys.argv[6]
wbN1 = sys.argv[7]
sN1 = sys.argv[8]
coln1 = int(sys.argv[9])
ncol1 = int(sys.argv[10])
bdate = sys.argv[11]


cowslist.fpyext_cwslst_at_base_date( wbN0, sN0, coln0, ncol0, index, name, 
                    wbN1, sN1, coln1, ncol1, bdate )

print( wbN1+ "/"  + sN1 + "のcowslistの基準日における 所属牛をsheet cowslistyyyymmddに抽出しました。")
