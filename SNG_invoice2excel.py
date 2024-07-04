from tkinter.filedialog import askopenfilename
import re, os
import pandas as pd

file_path = askopenfilename(initialdir='.')

invoiceNo = re.search('\d{10}', file_path).group()

try:
    os.mkdir('./'+invoiceNo)
except:
    print('Directory already exist.')

df = pd.read_excel(file_path)

# Create receiving file.
df_receiving = df[['Barcode', 'QtyShipped', 'UnitPrice']]
df_receiving.columns=['Code', 'Qty on Ord', 'Cost']

df_receiving.to_csv(f'{invoiceNo}/{invoiceNo}_receiving.txt', sep='\t',index=False)

# Create new item file.
df_new_item = df[['Barcode']]
df_new_item.insert(loc=len(df_new_item.columns),column='Description', value=df['StyleID'].str.cat(' '+df['Color']))
df_new_item.insert(loc=len(df_new_item.columns), column='Ext Desc', value=df['Description'].str.cat(' #'+df['Color']))
df_new_item.insert(loc=len(df_new_item.columns), column='Cost', value=df['UnitPrice'])
df_new_item.insert(loc=len(df_new_item.columns), column='Regular Price', value=0)
df_new_item.insert(loc=len(df_new_item.columns), column='Department', value=pd.NA)

df_new_item.to_csv(f'{invoiceNo}/{invoiceNo}_new_item.txt', sep='\t', index=False)

# Create supplier cost file.
df_supplier = df[['Barcode']]
df_supplier.insert(loc=len(df_supplier.columns),column='Cost', value=df['UnitPrice'])
df_supplier.insert(loc=len(df_supplier.columns),column='ReorderNo', value=df['StyleID'])
df_supplier.insert(loc=len(df_supplier.columns),column='MinimumOrder', value=0)
df_supplier.insert(loc=len(df_supplier.columns),column='MPQ', value=0)

df_supplier.to_csv(f'{invoiceNo}/{invoiceNo}_supplier.txt', sep='\t', index=False)