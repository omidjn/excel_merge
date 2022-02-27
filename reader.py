import xlsxwriter
import threading
import openpyxl
import os

def main():
    threading.Timer(60 * 5, main).start()
    all_data = []

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
    
    home = os.path.expanduser('~')
    dir = os.path.join(home, 'Documents')
    file_dir = os.path.join(dir, "files")

    files = os.listdir(file_dir)

    for i in files:
        if i.endswith(".xlsx"):
            all_data.append(read(os.path.join(file_dir, i)))

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


if __name__ == '__main__':
    main()
