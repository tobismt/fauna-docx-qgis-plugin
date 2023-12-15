import os
import pandas as pd
from openpyxl import load_workbook

os.chdir(r"C:\Users\Uni\Desktop\Arten")
df = pd.DataFrame()

def get_data(path):
    df = pd.read_excel(path)
    return df


for file in os.listdir():
    if file.endswith('.xlsx'):
        dd = get_data(file)
        df = pd.concat([df, dd])

df.to_csv('fauna.csv', sep='|')