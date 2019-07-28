import requests
import zipfile
import pandas as pd
import glob
import os
import xlrd

# Корневая папка
root = os.path.dirname(os.path.abspath(__file__))

archives_directory = ''.join([root, '\\archives'])
excels_directory = ''.join([root, '\\excels'])
csv_directory = ''.join([root, '\\csv'])

# ID регионов
ids = [str(id) for id in [	'ESTAT086117',
							'ESTAT086118',
							'ESTAT086119',
							'ESTAT086120',
							'ESTAT086121',
							'ESTAT086123',
							'ESTAT086124',
							'ESTAT086125',
							'ESTAT086126',
							'ESTAT086127',
							'ESTAT086128',
							'ESTAT086129',
							'ESTAT086131',
							'ESTAT086132',
							'ESTAT086133',
							'ESTAT086134',
							'ESTAT267724']]

def download_data():
    if not os.path.exists(archives_directory):
        os.makedirs(archives_directory)

    print ('\nDownloading data ...')
    for i in ids:
        url = ''.join(['http://stat.gov.kz/api/getFile/?docId=', i])
        print (url)
        file = ''.join([archives_directory, '\\', i, '.zip'])

        ufr = requests.get(url)
        f = open(file, 'wb')
        f.write(ufr.content)
        f.close()

def unzip_data():
    if not os.path.exists(excels_directory):
        os.makedirs(excels_directory)

    print ('\nUnziping ...')
    for i in ids:
        file = ''.join([archives_directory, '\\', i, '.zip'])
        print (file)
        zip_ref = zipfile.ZipFile(file, 'r')
        zip_ref.extractall(excels_directory)
        zip_ref.close()

def parse_data():
    if not os.path.exists(csv_directory):
        os.makedirs(csv_directory)

    wb_pattern = os.path.join(excels_directory, '*.xls')
    excels = glob.glob(wb_pattern)
    all_data = pd.DataFrame()

    print ('\nParsing excels ...')
    for f in excels:
        print (f)
        ef = xlrd.open_workbook(f)
        cnt_sh = len(ef.sheet_names())
        for cnt in list(range(0,(cnt_sh))):
            df = pd.read_excel(f, sheet_name = cnt, index_col=None, skiprows=3)
            all_data = all_data.append(df, ignore_index=True)

    all_data.to_csv(''.join([csv_directory, '\\', 'bin.csv']), sep=';', encoding = 'utf8')
    all_data.to_excel(''.join([csv_directory, '\\', 'bin.xlsx']))
    
download_data()
unzip_data()
parse_data()