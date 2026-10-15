import pandas as pd 
import openpyxl

from pathlib import Path





"""
first lets pull the files and distplat the first three row 

"""

BASE = Path(__file__).resolve().parents[1]

input_path = BASE / "data" / "extracted"

list_of_files = [file for file in input_path.glob("*.xlsx")]

f = list_of_files[0]



list_of_columns = []

df = pd.read_excel( f , nrows= 0, engine= "openpyxl")

df = df.columns.to_list()
print(df)
list_of_columns.append(df)
df = pd.read_excel( f , nrows= 0, header= 1 , engine= "openpyxl")

df = df.columns.to_list()
print(df)
list_of_columns.append(df)
df = pd.read_excel( f , nrows= 0, header=2,engine= "openpyxl")

df = df.columns.to_list()
print(df)
list_of_columns.append(df)
df = pd.read_excel( f , nrows= 0, header= 3 ,  engine= "openpyxl")

df = df.columns.to_list()
print(df)
list_of_columns.append(df)
print(list_of_columns)


