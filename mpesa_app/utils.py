import pikepdf
from copy import copy
import PyPDF2
import io
import re
from io import StringIO, BytesIO
import re
from PyPDF2 import PdfFileReader
from openpyxl import Workbook
from openpyxl.styles import Font, Color, colors
import random, string
from openpyxl.writer.excel import save_virtual_workbook
import openpyxl
from openpyxl import load_workbook
import os
import pandas as pd
import pandas.io.formats.excel
pandas.io.formats.excel.header_style = None
import xlsxwriter

i = 0
#regex pattern for mpesa statement
regex1 = r'(C.+ Name)(.+)(M.+ Number)(\d+)(E\w+ Address)(.+)(D[ a-zA-Z]+Statement)(.+)(S.+ Period)(.+ - \d{2} \w+ \d{4})'
regex = r'(\w{10})(\d{4}-\d{2}-\d{2} \d{2}\:\d{2}\:\d{2})(.+?)(Completed)(.*?\.\d{2})(.*?\.\d{2})'

def extract_from_pdf(file, password):
    #decrypting the encrypted pdf file
    content = pikepdf.open(file, password=password)
    inmemory_file = BytesIO()
    content.save(inmemory_file)
    #reading and extracting data from the decrypted pdf file
    pdf_reader = PyPDF2.PdfFileReader(inmemory_file)
    num_pages = pdf_reader.getNumPages()

    extracted_data = StringIO()
    for page in range(num_pages):
        extracted_data.writelines(pdf_reader.getPage(page).extractText())

    return num_pages, extracted_data


def random_str(length=8):
	s = ''
	for i in range(length):
		s += random.choice(string.ascii_letters + string.digits)

	return s


def parse_mpesa_content(extracted_data):
    extracted_data.seek(0)
    lines = extracted_data.read()
    matches = re.compile(regex).findall(lines)
    matches2 = re.compile(regex1).findall(lines)

    fb = Font(name='Calibri', color=colors.BLACK, bold=True, size=11, underline='single')
    i = 0
    #creating the spreadheet
    book = Workbook()
	# grab the active worksheet
    sheet = book.active
	#excel styling 2
    ft = Font(name='Calibri', color=colors.BLUE, bold=True, size=11, underline='single')

    sheet['A1'] = 'RECEIPT NO'
    sheet['B1'] = 'COMPLETION TIME'
    sheet['C1'] = 'DETAILS'
    sheet['D1'] = 'TRANSACTION STATUS'
    sheet['E1'] = 'VALUE'
    sheet['F1'] = 'BALANCE'

    a1 = sheet['A1']
    b1 = sheet['B1']
    c1 = sheet['C1']
    d1 = sheet['D1']
    e1 = sheet['E1']
    f1 = sheet['F1']

    a1.font = ft
    b1.font = ft
    c1.font = ft
    d1.font = ft
    e1.font = ft
    f1.font = ft


	#adding every match to the excel file
    while i < len(matches):
	    # print(matches[i])
        sheet.append(matches[i])
        i = i + 1

    filename = random_str()
    book.save(filename)
    f = open(filename, 'rb')
    file = BytesIO(f.read())
    f.close()
    os.remove(filename)

    return file, matches2

def find_name(matches2):
    for match in matches2:
        print(match[1])
        
    return match[1]

def paidin(workbook):
    excel_df = pd.read_excel(workbook)
    excel_df['VALUE'] = excel_df['VALUE'].astype(str).str.replace(',', '').astype(float)
    paidinall = excel_df[excel_df['VALUE']>0]
    # paidinall.set_index('DETAILS', inplace=True)
    paidin = paidinall[['VALUE', 'DETAILS']].sort_values('DETAILS').groupby(['DETAILS'], as_index=False)['VALUE'].sum()
    def format(row):
        index = None
        reg = re.search(r'\d', row['DETAILS'])
        if reg:
            index = reg.start()
        row['DETAILS'] = row['DETAILS'][:index]
        return row
    idx = paidin.apply(format, axis=1).groupby(['DETAILS'], as_index=False).apply(lambda r: r).sort_values(['DETAILS', 'VALUE'], ascending=False).index
    paidin = paidin.loc[idx]
    subtotal = paidin['VALUE'].sum()
    excel_df2 = pd.DataFrame({'VALUE':[subtotal], 'DETAILS': 'Grand Total'})
    df_append = paidin.append(excel_df2, ignore_index=False)
    df_append.rename(columns={'VALUE':'AMOUNT'}, inplace=True)

    return df_append

def withdrawal(workbook):
    excel_df = pd.read_excel(workbook)
    excel_df['VALUE'] = excel_df['VALUE'].astype(str).str.replace(',', '').astype(float)
    withdrawal = excel_df[excel_df['VALUE']<0]
    # withdrawal.set_index('DETAILS', inplace=True)
    withdrawn = withdrawal[['VALUE', 'DETAILS']].sort_values('DETAILS').groupby(['DETAILS'], as_index=False)['VALUE'].sum()
    withdrawn['VALUE'] = withdrawn['VALUE'].astype(str).str.replace('-', '').astype(float)
    def format(row):
        index = None
        reg = re.search(r'\d', row['DETAILS'])
        if reg:
            index = reg.start()
        row['DETAILS'] = row['DETAILS'][:index]
        return row
    idx = withdrawn.apply(format, axis=1).groupby(['DETAILS'], as_index=False).apply(lambda r: r).sort_values(['DETAILS', 'VALUE'], ascending=False).index
    withdrawn = withdrawn.loc[idx]
    subtotal = withdrawn['VALUE'].sum()
    excel_df2 = pd.DataFrame({'VALUE':[subtotal], 'DETAILS': 'Grand Total'})
    df_append = withdrawn.append(excel_df2, ignore_index=False)
    df_append.rename(columns={'VALUE':'AMOUNT'}, inplace=True)

    return df_append

def listing(paidin, withdrawn):
    dfs = [paidin, withdrawn]

    return dfs


def dfs_tabs(df_list, sheet_list, file_name):
    file_name = BytesIO()
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    workbook = writer.book
    fmt = workbook.add_format({'align':'left', 'size':10, 'font_name': 'Times New Roman'})
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, sheet_name=sheet, startrow=0, startcol=0, index=False)
    worksheet = writer.sheets['PAID IN DATA']
    worksheet2 = writer.sheets['WITHDRAWN DATA']
    worksheet.set_column(0, 2, 90.0, fmt)
    worksheet2.set_column(0, 2, 90.0, fmt)
    writer.save()

    return file_name
