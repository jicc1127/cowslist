# cowslist
PS C:\Users\inoue\Dropbox\rep\cowslist>  python ps_cowslistmanual00_args.py
-----cowslist Manual00---------------------------------------------------v2.1------
1.input individual informations from cowshistory's data to cowslist
PS> python ps_fpyind_inf_to_cowslist_args.py wbN0 sheetN0 colidno0 wbN1 sheetN1 colidno1
wbN0 : AB_cowshistory.xlsx, sheetN0 : ABFarm, colidno0 : 2 (column number fo idno0),
wbN1 : AB_cowslist.xlsx, sheetN1 : cowslist, colidno1 : 2 (column number fo idno1)

2.input individual transfer informations from cowshistory's data to cowslist
PS> python ps_fpyind_trsinf_to_cowslist_args.py wbN0 sheetN0 colidno0
wbN1 sheetN1 colidno1 name
wbN0 : AB_cowshistory.xlsx, sheetN0 : ABFarm, colidno0 : 2 (column number fo idno0),
wbN1 : AB_cowslist.xlsx, sheetN1 : cowslist, colidno1 : 2 (column number fo idno1)
name : 氏名または名称

3.modify cowslist data from a new cowslist
PS> python ps_fpymod_cowslist_args.py wbN sheetN snewN ncol
wbN: ..\AB_cowslist.xlsx, sheetN: cowslist, scolN: cowslist_new, ncol: 20
新しい情報に変更があれば修正する

4.register new cows from a new cowslist
PS> python ps_fpyreg_newcows_args.py wbN sheetN snewN ncol
wbN: ..\AB_cowslist.xlsx, sheetN: cowslist, scolN: cowslist_new, ncol: 20
新規牛をExcelsheet cowslist追加

5.make an ExcelSheet if it dose not exist
PS> python ps_fpymkxlsheet_args.py wbN sheetN scolN r
wbN: ..\AB_cowslist.xlsx, sheetN: ABFarmout, scolN: columns, r: 1
2枚のsheet ABFarmin, ABFarmout がなければ、作成する

6.separate move-out cows from move-in in a cowslist
PS> python ps_fpysepclst_outfrmin_args.py wbN sheetN ncol index
wbN: ..\AB_cowslist.xlsx, sheetN: ABFarm, ncol: 20, r: index : 15
検索日における所属牛（転入牛move-in)と転出牛(move-out)の情報に分け、
2枚のsheet ABFarmin, ABFarmout を作成する
---------------------------------------------------------2024/2/11 by jicc---------

AB_cowslist.xlsx/cowslist : ABFarmのcowslist

- sheet cowslist のcolumnを以下のように設定する :
| cowslist_id | cowidNo | eartagNo | DHITNo | name | breed | birthday | sex |
| sire_code | sire_name | damidNo | dameartagNo | in_date | in_reason | from | out_date |
| out_reason | to | note | base_date |
or
| cowslist_id | 個体識別番号 | 耳標 | 検定番号 | 名前 | 品種 | 生年月日 | 性別 |
| 種雄牛コード | 種雄牛名 | 母牛個体識別番号 | 母牛耳標 | 導入年月日 | 導入理由 | 導入元 | 転出年月日 |
| 転出理由 | 転出先 | note | 基準日 |
- sheet cowlist の2列　cowidNo(個体識別番号) に　ABFarm 所属牛の個体識別番号をリストする。
- cowshistory の CowsHistory_webscrsysによって　web検索によりAB_cowshistory.xlsx/ABFarm にcowshistoryを作成する。
    
    sheet ABFarm
| LineNo | 個体識別番号 | 出生の年月日 | 雌雄の別 | 母牛の個体識別番号 | 種別 |
| No | 異動内容 | 異動年月日 | 住所 | 氏名または名称 | 検索年月日 |

- cowslist Manual00
    1. AB_cowshistory.xlsx/ABFarm より AB_cowslist.xlsx/cowslistに個体情報を入力する。
    2. AB_cowshistory.xlsx/ABFarm より AB_cowslist.xlsx/cowslistに異動情報を入力する。
- 新しい AB_cowslist.xlsx/cowslistyyyymmdd に個体情報を入力 cowslist Manual00 1.
    
    web情報以外で、個体情報に変更があれば、
    
    cowslist Manual00 3. で、AB_cowslist.xlsx/cowslist (original list)を変更する。
    
    ABFarm の個体情報で、分娩後 DHITNo(検定番号) 取得するため、個体情報が変更されるときに対応した。
    
- cowslist Manual00
    1. 新規情報AB_cowslist.xlsx/cowslistyyyymmdd で、
        
        新規牛をAB_cowslist.xlsx/cowslist (original list) に追加入力する。
        
    
    2. AB_cowshistory.xlsx/ABFarm より AB_cowslist.xlsx/cowslistに異動情報を入力する。
    
    この作業で、base_date(基準日)に最新検索日の日付が入る。
