# Wrapper for xlswriter as pandas to_excel uses large amount of memory
import xlsxwriter
import numpy as np
from datetime import datetime


class Excel:
    def __init__(self, name):
        self.workbook = xlsxwriter.Workbook(name, {'constant_memory': True})

    def addSheet(self, sheetName, df):
        print('Writing:', sheetName, datetime.now().strftime("%H:%M:%S"))
        bold = self.workbook.add_format({'bold': True})
        worksheet = self.workbook.add_worksheet(sheetName)
        df = df.replace({np.nan: None})
        worksheet.write_row(0, 0, list(df.columns), bold)
        row = 1
        for r in df.values.tolist():
            worksheet.write_row(row, 0, r)
            row += 1

    def save(self):
        self.workbook.close()


class Excel2:
    def __init__(self, name):
        self.workbook = xlsxwriter.Workbook(name, {'constant_memory': True})

    def addSheet(self, sheetName, df):
        print('Writing:', sheetName, datetime.now().strftime("%H:%M:%S"))
        bold = self.workbook.add_format({'bold': True})
        worksheet = self.workbook.add_worksheet(sheetName)
        df = df.replace({np.nan: None})
        col = 0
        columns = list(df.columns)

        for column in columns:
            worksheet.write(0, col, column, bold)
            col += 1

        row = 1
        for r in df.to_dict('records'):
            col = 0
            for column in columns:
                worksheet.write(row, col, r[column])
                col += 1
            row += 1

    def save(self):
        self.workbook.close()
