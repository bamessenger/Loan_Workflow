import pandas as pd
import pathlib

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


class XLFile:

    def fileRead(self, encompPath):
        fileExt = pathlib.Path(encompPath).suffix
        if fileExt == '.csv':
            self.encmpData = pd.read_csv(encompPath, header=0)
        else:
            self.encmpData = pd.read_excel(encompPath, engine='openpyxl')

    def excelWrite(self, wrkflwPath):
        wrkbk = load_workbook(wrkflwPath)
        delSheet = wrkbk.get_sheet_by_name('tblEncompass')
        wrkbk.remove_sheet(delSheet)
        writer = pd.ExcelWriter(wrkflwPath, mode='a',
                                datetime_format='MM-DD-YYYY', engine='openpyxl')
        writer.book = wrkbk
        self.encmpData.to_excel(writer, sheet_name='tblEncompass', startcol=1,
                                index=False)
        sheet = wrkbk.get_sheet_by_name('tblEncompass')
        table = Table(displayName='tblEncompass',
                      ref='B1:' + get_column_letter(sheet.max_column) + str(
                          sheet.max_row))
        style = TableStyleInfo(name='TableStyleMedium11',showRowStripes=False,
                               showColumnStripes=False)
        table.tableStyleInfo = style
        sheet.add_table(table)
        writer.save()
        writer.close()

