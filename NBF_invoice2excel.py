from tkinter.filedialog import askopenfilename
import PyPDF2, re, os
import pandas as pd

# root_path = "Z:/excel files/00 RMH Sale report/"
root_path = ''

# Define the columns
columns = [
    'Item Code',
    'Description',
    'Qty Order',
    'BackOrder',
    'Shipped',
    'Unit Price',
    'Total'
]

# Function to extract text from a PDF file
def extract_text_from_pdf(file_path):
    pdf_text = ""
    invoiceNo = ""
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text = page.extract_text()
            if invoiceNo == "":
                invoiceNo = re.search(r'PSI\d{5}', text).group()
            if page_num == len(reader.pages)-1:
                pdf_text += text[67:text.find('Remark')]
            elif page_num == 0:
                pdf_text += text[67:text.find('SOLD TO')]
            else:
                pdf_text += text[67:text.find('Printed Date')]
    return pdf_text, invoiceNo

# Function to parse the extracted text and extract invoice data
def parse_invoice_data(pdf_text):
    lines = pdf_text.split("\n")
    cursor = 0
    invoice_data = []
    row_data = []
    for line in lines:
        if cursor%len(columns) == 0 and row_data != []:
            invoice_data.append(row_data)
            row_data = []

        if cursor%len(columns) == 2 and not line.isnumeric():
            row_data[-1] += line
        elif cursor%len(columns) == 4:
            if '$' in line:
                row_data.append('0')
                cursor += 1
            row_data.append(line)
            cursor += 1
        else:
            row_data.append(line)
            cursor += 1
    return invoice_data


# Extract text from the PDF
file_path = askopenfilename(initialdir='.')
pdf_text, invoiceNo = extract_text_from_pdf(file_path)

try:
    os.mkdir('./'+invoiceNo)
except:
    print(f"{invoiceNo} already exist.")

# Parse the extracted text to get invoice data
invoice_data = parse_invoice_data(pdf_text)

# Create a DataFrame
df = pd.DataFrame(invoice_data, columns=columns)
df['Item Code'] = df['Item Code'].str.slice(stop=18)

# Load NBF(Chade) inventory list.
NBF_list = pd.read_excel(root_path+'inv_data/NBF_Chade Fashions.xlsx', dtype=str)
NBF_list['Item Code'] = NBF_list['No.'].str.slice(stop=18)

df = df.merge(NBF_list[['Item Code', 'No.', 'UPC Code']], how='left')

df['Item'] = pd.NA
df['Color'] = pd.NA

for row in df.iterrows():
    print(row[1]['No.'])
    Item, Color = row[1]['No.'].split('-')
    df.loc[row[0], 'Item'] = Item
    df.loc[row[0], 'Color'] = Color

# Create receiving file.
df_receiving = df[['UPC Code', 'Shipped', 'Unit Price']]
df_receiving.columns=['Code', 'Qty on Ord', 'Cost']
df_receiving.loc[:, 'Cost'] = df_receiving['Cost'].str.strip('$ ')
df_receiving = df_receiving[df_receiving['Qty on Ord']!='0'].reset_index()

df_receiving.to_csv(f'{invoiceNo}/{invoiceNo}_receiving.txt', sep='\t', index=False)

# Create new item file.
df_new_item = df[['UPC Code']]
df_new_item.columns = ['Barcode']
df_new_item.insert(loc=len(df_new_item.columns),column='Description', value=df['Item'].str.cat(' '+df['Color']))
df_new_item.insert(loc=len(df_new_item.columns), column='Ext Desc', value=df['Description'].str.cat(' #'+df['Color']))
df_new_item.insert(loc=len(df_new_item.columns), column='Cost', value=df['Unit Price'].str.strip('$ '))
df_new_item.insert(loc=len(df_new_item.columns), column='Regular Price', value=0)
df_new_item.insert(loc=len(df_new_item.columns), column='Department', value=pd.NA)

df_new_item.to_csv(f'{invoiceNo}/{invoiceNo}_new_item.txt', sep='\t', index=False)

# Create supplier cost file.
df_supplier = df[['UPC Code']]
df_supplier.columns = ['Barcode']
df_supplier.insert(loc=len(df_supplier.columns),column='Cost', value=df['Unit Price'])
df_supplier.insert(loc=len(df_supplier.columns),column='ReorderNo', value=df['Item'])
df_supplier.insert(loc=len(df_supplier.columns),column='MinimumOrder', value=0)
df_supplier.insert(loc=len(df_supplier.columns),column='MPQ', value=0)

df_supplier.to_csv(f'{invoiceNo}/{invoiceNo}_supplier.txt', sep='\t', index=False)