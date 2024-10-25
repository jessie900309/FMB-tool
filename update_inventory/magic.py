import pandas as pd
from pandas import DataFrame
from typing import List

# Define file path and sheet names
file_path: str = '庫存.xlsx'
original_sheet_name: str = input("舊的工作表名稱：")
update_sheet_name: str = input("新的工作表名稱：")
new_sheet_name: str = '更新結果'

# Load data from Excel sheets into DataFrames
original_data: DataFrame = pd.read_excel(file_path, sheet_name=original_sheet_name)
update_data: DataFrame = pd.read_excel(file_path, sheet_name=update_sheet_name)

# Columns to clean and use for creating 'product_id'
columns_to_clean: List[str] = ['貨品編號', '寛度', '長度']

# Fill NaN values with empty strings and ensure columns are of type str
for col in columns_to_clean:
    original_data[col] = original_data[col].fillna('').astype(str)
    update_data[col] = update_data[col].fillna('').astype(str)

# Fill remaining NaN values in original_data with empty strings
original_data = original_data.fillna('')

# Create 'product_id' by concatenating specified columns with '-'
original_data['product_id'] = original_data[columns_to_clean].apply('-'.join, axis=1)
update_data['product_id'] = update_data[columns_to_clean].apply('-'.join, axis=1)

# Set 'product_id' as index for both DataFrames, keeping it as a column
original_data.set_index('product_id', drop=False, inplace=True)
update_data.set_index('product_id', drop=False, inplace=True)

# Update '庫存粒數' in original_data with values from update_data where 'product_id's match
original_data.update(update_data[['庫存粒數']])

# Set '庫存粒數' to 0 in original_data for products not present in update_data
products_to_zero = original_data.index.difference(update_data.index)
original_data.loc[products_to_zero, '庫存粒數'] = 0

# Identify new products in update_data not present in original_data
new_products = update_data.index.difference(original_data.index)
rows_to_append: DataFrame = update_data.loc[new_products]

# Append new products to original_data
original_data = pd.concat([original_data, rows_to_append], ignore_index=True)
original_data = original_data.drop('product_id', axis='columns')

# Write the final DataFrame to a new sheet in the Excel file
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
    original_data.to_excel(writer, sheet_name=new_sheet_name, index=False)

print("花花魔法球完成更新囉~ OuO!!!")
# print(original_data)
