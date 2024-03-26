from tkinter.filedialog import askopenfilename
import pandas as pd

path = askopenfilename(initialdir='.')

df = pd.read_html(path, displayed_only=False)[0]

invoiceNo = df['InvoiceCode'][0]
df_receiving = df[['Barcode', 'Qty', 'Price']]
df_receiving.columns=['Code', 'Qty on Ord', 'Cost']
df_receiving.loc[:, 'Cost'] = df_receiving['Cost'].str.strip('$ ')

df_receiving.to_csv(f'{invoiceNo}_receiving.csv', index=False)

df_new_item = df[['Barcode']]
df_new_item.insert(loc=len(df_new_item.columns),column='Description', value=df['Item'].str.cat(' '+df['Color']))
df_new_item.insert(loc=len(df_new_item.columns), column='Ext Desc', value=df['Name'].str.cat(' #'+df['Color']))
df_new_item.insert(loc=len(df_new_item.columns), column='Cost', value=df['Price'].str.strip('$ '))
df_new_item.insert(loc=len(df_new_item.columns), column='Regular Price', value=0)
df_new_item.insert(loc=len(df_new_item.columns), column='Department', value=pd.NA)

df_new_item.to_csv(f'{invoiceNo}_new_item.csv', index=False)

df_supplier = df[['Barcode']]
df_supplier.insert(loc=len(df_supplier.columns),column='Cost', value=df['Price'])
df_supplier.insert(loc=len(df_supplier.columns),column='ReorderNo', value=df['Item'])
df_supplier.insert(loc=len(df_supplier.columns),column='MinimumOrder', value=0)
df_supplier.insert(loc=len(df_supplier.columns),column='MPQ', value=0)

df_supplier.to_csv(f'{invoiceNo}_supplier.csv', index=False)