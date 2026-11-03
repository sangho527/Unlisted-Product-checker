import config as cfg
import pandas as pd, re
s=pd.read_csv(cfg.SELLMATE_STOCK,encoding='cp949',dtype=str)
s=s.loc[:,~s.columns.str.startswith('Unnamed')]
for c in s.columns: s[c]=s[c].astype(str).str.replace(r'^=?"*|"*$','',regex=True).str.replace('=','').str.strip('"')
import os
x=pd.ExcelFile(cfg.MUSINSA_ALL) if cfg.MUSINSA_ALL and os.path.exists(cfg.MUSINSA_ALL) else None   # 등록 현황 점검(step1)에서만 사용
