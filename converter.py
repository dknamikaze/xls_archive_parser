from bs4 import BeautifulSoup
from lxml import html
import pandas as pd
import openpyxl
import re,os


def archive_reader(file_path):
    print(file_path)
    with open(file_path, 'rb') as f:
        data = f.read()
    # print(data)
    return data

pattern = r'^\d{2}-[A-Za-z]{3}-\d{4}$'
def check_date_format(date_string):
    if re.match(pattern, date_string):
        return True
    return False

# def read_xlsb_rows(file_path):
#     """
#     Generator function to read rows from an .xlsb file (Excel Binary Workbook format).
#     """
#     with open(file_path, 'rb') as file:
#     # Read the file in chunks (e.g., 1024 bytes at a time)
#         while chunk := file.read(1004):
#             yield chunk

# # Display the first few bytes to inspect the structure
# for chunk in read_xlsb_rows('nsedata\Archive_Data_Oct24.xls'):
#     print(chunk)
#     break

months = ['Aug','Sep','Oct','Nov','Dec']
output_excel = "output_fii_stock.xlsx"

output_dict = {
    "Equity":{},
    "Index Futures":[],
    "Index Options":[],
    "Stock Futures":[],
    "Stock Options":[]
}

for month in months:
    # month='Sep'
    data = archive_reader(f"nsedata\Archive_Data_{month}24.xls")
    data_xml = html.fromstring(data)
    print(data_xml)

    tables = data_xml.xpath('./table')
    # print(tables[0].xpath('.//text()'))

    for row in tables[0].xpath('./tbody/tr'):
        row = row.xpath(".//text()")
        # print(row)
        if row:
            if check_date_format(row[0]):
                if row[6][-1]==')':
                    row[6] = f'-{row[6][1:-1]}'
                output_dict["Equity"][row[0]]=row[6]
    
    date = ''
    for row in tables[1].xpath('./tbody/tr'):
        row = row.xpath(".//text()")
        # print(row)
        if row:
            if check_date_format(row[0]):
                date = row[0]
                output_dict[row[1]].append([date]+row[2:])
            else:
                if date and row[0] in output_dict.keys():
                    output_dict[row[0]].append([date]+row[1:])
            
    # break


print(output_dict)


with pd.ExcelWriter(output_excel, mode='w', engine='openpyxl') as writer:
    for key in output_dict.keys():
        stock_data = output_dict[key]
        if key == "Equity":
            print(key)
            df = pd.DataFrame(list(stock_data.items()), columns=['Date', 'FII Net Trade'])
        else:
            print(key)
            df =  pd.DataFrame(list(stock_data), columns=['Date', "No. of Buy Contracts","Amount in Crore"	,"No. of Sell Contracts","Amount in Crore",	"No. of OI Contracts","Amount in Crore"])

        # Write to Excel
        df.to_excel(writer, sheet_name=key, index=False)

# for row in tables[1].xpath('./tbody/tr'):
#     row = row.xpath(".//text()")
#     print(row)
    
