import glob
import pandas as pd
import openpyxl

files=glob.glob('アンケート*.xlsx')

list=[]

for file in files:
    list.append(pd.read_excel(file))