import pandas as pd
import pathlib

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.utils import get_column_letter


class XLFile:

    def fileRead(self, encompPath):
        fileExt = pathlib.Path(encompPath).suffix
        if fileExt == '.csv':
            self.encmpData = pd.read_csv(encompPath, header=0)
        else:
            self.encmpDataAll = pd.read_excel(encompPath, engine='openpyxl')
            self.encmpDataAll.columns = \
                self.encmpDataAll.columns.str.replace(' ','')
            self.encmpDataOpen = self.encmpDataAll[
                self.encmpDataAll.ClosedDate.isnull()]
            self.encmpDataOpen = self.encmpDataOpen.reset_index()

    def excelWrite(self, wrkflwPath):
        wrkbk = load_workbook(wrkflwPath)
        delSheetAll = wrkbk.get_sheet_by_name('tblEncompassAllAct')
        delSheetOpen = wrkbk.get_sheet_by_name('tblEncompassOpen')
        wrkbk.remove_sheet(delSheetAll)
        wrkbk.remove_sheet(delSheetOpen)
        writer = pd.ExcelWriter(wrkflwPath, mode='a',
                                datetime_format='MM-DD-YYYY', engine='openpyxl')
        writer.book = wrkbk
        self.encmpDataAll.to_excel(writer, sheet_name='tblEncompassAllAct',
                                   startcol=1, index=False)
        sheet = wrkbk.get_sheet_by_name('tblEncompassAllAct')
        table = Table(displayName='tblEncompassAllAct',
                      ref='B1:' + get_column_letter(sheet.max_column) + str(
                          sheet.max_row))
        style = TableStyleInfo(name='TableStyleMedium11', showRowStripes=False,
                               showColumnStripes=False)
        table.tableStyleInfo = style
        sheet.add_table(table)
        self.encmpDataOpen.sort_values(by=[
            'ApplicationDate'], inplace=True, ignore_index=True)
        self.encmpDataOpen.to_excel(writer, sheet_name='tblEncompassOpen',
                                   startcol=0, index=True)
        sheet = wrkbk.get_sheet_by_name('tblEncompassOpen')
        table = Table(displayName='tblEncompassOpen',
                      ref='B1:' + get_column_letter(sheet.max_column) + str(
                          sheet.max_row))
        style = TableStyleInfo(name='TableStyleMedium11', showRowStripes=False,
                               showColumnStripes=False)
        table.tableStyleInfo = style
        sheet.add_table(table)
        writer.save()
        writer.close()
