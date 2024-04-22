import pandas as pd
from tkinter import simpledialog, messagebox
import json, os

# root_path = "Z:/excel files/00 RMH Sale report/"
root_path = ''
invoiceNo = simpledialog.askstring("Boyang Invoice to Excel", "Input invoice No.\t\t\t\t\t")

with open(root_path+"appdata/by_address.json") as f:
    temp = json.load(f)
    BY_INVOICE_ADDRESS = temp['BY_INVOICE_ADDRESS']

df = pd.read_html(BY_INVOICE_ADDRESS+invoiceNo)
# try:
#     df = pd.read_html(BY_INVOICE_ADDRESS+invoiceNo)
# except Exception as e:
#     messagebox.showerror("Error", f"Wrong invoice Number \n{e}")
#     raise SystemExit

# Select item table from invoice
itemList = df[9].drop(columns=[4,5], axis=1)

# Preprocessing
itemList.columns = itemList.iloc[0]
itemList = itemList[1:].reset_index(drop=True)
itemList['DESCRIPTION'] = itemList['DESCRIPTION'].str.split(',')

# Select required columns
new_list = pd.DataFrame(columns=['Item', 'QTY', 'Color', 'Price'])

# Expand items list
for i in range(len(itemList)):
    for des in itemList['DESCRIPTION'][i]:
        des = des.strip()
        if des == '#1' or des == '#2':
            continue
        else:
            color, qty = des.split(':')
            new_list.loc[len(new_list)] = [itemList['ITEM'][i], qty, color.strip('#'), itemList['UNIT PRICE'][i]]

# Drop backordered items
new_list['QTY'] = new_list['QTY'].str.replace('\(.\)', '', regex=True)
new_list.drop(new_list[new_list['QTY']=='0'].index, inplace=True)

# Load BY inventory list
BY_list = pd.read_excel(root_path+'inv_data/BY_InventoryListAll.xls', skiprows=3)

# Preprocessing
BY_list = BY_list[['Item Name', 'Color', 'Barcode']]
BY_list['Item'] = BY_list['Item Name'].str.replace(' ', '')
BY_list['Item'] = BY_list['Item'].str.replace('-', '')
BY_list['Item'] = BY_list['Item'].str.replace('TOP', 'TP')

new_list = new_list.merge(BY_list, how='left', left_on=['Item', 'Color'], right_on=['Item', 'Color'])

try:
    os.mkdir('./'+invoiceNo)
except:
    print(f'{invoiceNo} directory already exist!')

# receiving import file create
df_receiving = new_list[['Barcode', 'QTY', 'Price']]
df_receiving.columns=['Code', 'Qty on Ord', 'Cost']

df_receiving.to_csv(f'{invoiceNo}/{invoiceNo}_receiving.txt', sep='\t',index=False)

# new item import file create
df_new_item = new_list[['Barcode']]
df_new_item.insert(loc=len(df_new_item.columns),column='Description', value=new_list['Item Name'].str.cat(' '+new_list['Color']))
df_new_item.insert(loc=len(df_new_item.columns), column='Ext Desc', value=new_list['Item Name'].str.cat(' #'+new_list['Color']))
df_new_item.insert(loc=len(df_new_item.columns), column='Cost', value=new_list['Price'].str.strip('$ '))
df_new_item.insert(loc=len(df_new_item.columns), column='Regular Price', value=0)
df_new_item.insert(loc=len(df_new_item.columns), column='Department', value=pd.NA)

df_new_item.to_csv(f'{invoiceNo}/{invoiceNo}_new_item.txt', sep='\t', index=False)

# supplier import file create
df_supplier = new_list[['Barcode']]
df_supplier.insert(loc=len(df_supplier.columns),column='Cost', value=new_list['Price'])
df_supplier.insert(loc=len(df_supplier.columns),column='ReorderNo', value=new_list['Item Name'])
df_supplier.insert(loc=len(df_supplier.columns),column='MinimumOrder', value=0)
df_supplier.insert(loc=len(df_supplier.columns),column='MPQ', value=0)

df_supplier.to_csv(f'{invoiceNo}/{invoiceNo}_supplier.txt', sep='\t', index=False)