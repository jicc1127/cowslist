# -*- coding: utf-8 -*-
"""
fpystrdate_to_yyyymmdd : 
    change str date yyyy/mm/dd to str yyyymmdd
    v1.0
    2024/2/21
    @author: inoue
    
"""
def fpystrdate_to_yyyymmdd( date ):
    """
    change str date yyyy/mm/dd to str yyyymmdd

    Parameters
    ----------
    date : str
        yyyy/mm/dd
    
    Returns
    -------
    str : yyyymmdd

    """
    strd = date.split('/')
    #strdate yyyy/mm/dd を '/' で分離
    #strd = [yyyy, mm, dd]
    print(strd)
    
    yyyy = strd[0] #year yyyy
    
    lmm = len(strd[1]) #month mm の文字数
    if lmm == 1:
        mm = '0' + strd[1] #add '0' first
    else: #lmm == 2: 
        mm = strd[1] #without change
        
    ldd = len(strd[2])
    if ldd == 1:
        dd = '0' + strd[2]
    else: #ldd == 2:
        dd = strd[2]
    
    yyyymmdd = yyyy + mm + dd
    
    return yyyymmdd

yyyymmdd = fpystrdate_to_yyyymmdd( '2024/2/21' )
print(yyyymmdd)
