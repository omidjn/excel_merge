import openpyxl

def read(loc):
    wb_obj = openpyxl.load_workbook(loc)
    sheet = wb_obj.active
    data = []

    rows = list(sheet.rows)
    rows = rows[5:]

    for row in rows:
        r = [x.value if x.value != None else '' for x in row]
        del r[0]
        data.append(r)
        
    return data

data = read("files/sample.xlsx")

print(data)
