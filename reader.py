from unidecode import unidecode
import xlsxwriter
import openpyxl
import os

all_data = []

def read(loc):
    wb_obj = openpyxl.load_workbook(loc)
    sheet = wb_obj.active
    data = []

    rows = list(sheet.rows)
    factor_num = (int(unidecode(str(rows[0][14].value).split("-")[0])))
    rows = rows[5:]

    if factor_num in factor_nums:

        for row in rows:
            r = [x.value if x.value != None else '' for x in row]
            del r[0]
            data.append(r)
            
        return data
    else:
        return None

def read_main(loc):
    wb_obj = openpyxl.load_workbook(loc)
    sheet = wb_obj.active
    data = []

    cols = list(sheet.columns)[0]

    for row in cols:
        if row.value != None:
            data.append(row.value)
        
    return data

home = os.path.expanduser('~')
dir = os.path.join(home, 'Documents/files/')
file_dir = os.path.join(dir, "Factors/")

files = os.listdir(file_dir)

factor_nums = read_main(dir + 'main.xlsx')

for i in files:
    if i.endswith(".xlsx"):
        d = read(os.path.join(file_dir, i))
        if d != None:
            print(all_data)
            all_data.append(d)


workbook = xlsxwriter.Workbook(os.path.join(dir, 'all.xlsx'))
worksheet = workbook.add_worksheet()

row = 0
column = 0

for i in all_data:
    for j in i:
        column = 0
        for x in j:
            worksheet.write(row, column, x)
            print(x, row, column)
            column += 1
        row += 1

workbook.close()
print("done")
