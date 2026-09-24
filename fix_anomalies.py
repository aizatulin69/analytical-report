import pandas as pd


df = pd.read_csv("csv_pages/Warehouse_Logistics.csv")

# видалення нульових данних
df = df[(df != 0).all(axis=1)]

df.to_excel("csv_pages/Warehouse_Logistics.xlsx")