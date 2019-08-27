import xlsxwriter
import os
import exceltools


class NewWorkbook:
    def __enter__(self):
        return self

    def __init__(self, aimfilename, dirname="data"):
        self.workbook = xlsxwriter.Workbook("/".join(
            [os.getcwd(), dirname, aimfilename]))

    def insertsheet(self, sheetname):
        self.workbook.add_worksheet(sheetname)
        return self.workbook.get_worksheet_by_name(sheetname)

    def insertrow(self, sheet, rowstr, datas):
        sheet.write_row(rowstr, datas)

    def insertcol(self, sheet, colstr, datas):
        sheet.write_column(colstr, datas)

    def insertcell(self, sheet, rowid, colid, data):
        sheet.write(rowid - 1, colid, data)

    def insertformula(self, sheet, rowid, colid, fmlstr):
        posi = str(exceltools.getcolname(colid)) + str(rowid)
        sheet.write_formula(posi, fmlstr)

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.workbook.close()
