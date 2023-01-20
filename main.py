from openpyxl import load_workbook

if __name__ == '__main__':
    workbook = load_workbook(filename="Tarifs2023.xlsx")
    w = workbook
    print(w.sheetnames)

    s= w['AK98']
    print(s['A3'].value)
