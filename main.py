import xlwings as xw
from documents import Excel

if __name__ == '__main__':
    excel = Excel(r'D:\source-code\edit-realtime-excel\Book1.xlsx')
    excel.cell_active()
    excel.cell_set(44)
    excel.cell_up()
    excel.cell_set(45)
