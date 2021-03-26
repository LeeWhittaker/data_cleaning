import pandas as pd
import openpyxl
import argparse
import configparser

def get_options():

    my_parser = argparse.ArgumentParser(description='Clean excel spreadsheet')

    my_parser.add_argument('-o',
                           '--options',
                           metavar='inpath',
                           action='store',
                           type=str,
                           required=True,
                           help='path to the database')

    args =  my_parser.parse_args()
    options = args.options

    config = configparser.ConfigParser()

    config.read(options)
    
    return config


config = get_options()

infile = config['MAIN']['InputWorkbook']
try:
    outfile = config['MAIN']['OutputWorkbook']
except:
    outfile = infile
sheet = config['MAIN']['SheetName']
column = config['MAIN']['ColumnName']
regex_in = config['MAIN']['CurrentExpression']
ex_out = config['MAIN']['NewExpression']

df = pd.read_excel(infile, sheet_name=sheet)
new_col = df[column].replace(r'%s'%regex_in, r'%s'%ex_out, regex=True)

xfile = openpyxl.load_workbook(infile)
worksheet = xfile[sheet]

for cell in worksheet[1]:
    if cell.value==column:
        col_id = cell.column_letter
        
for i in range(len(new_col)):
    worksheet['%s%i' %(col_id, i+2)]=new_col[i]

xfile.save(outfile)
