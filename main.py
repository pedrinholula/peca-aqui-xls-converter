# -*- coding: utf-8 -*-
import pandas as pd

in_file = "in/LISTA JULHO VENDAS GERAL.xlsx"

df = pd.read_excel(in_file,heade=1)

print(df)
