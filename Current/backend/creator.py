import openpyxl as oxl
from backend.spreadsheet import *



class spreadsheet:
    def __init__(self, keys, catagories, name="Form") -> None:
        self.name = name 
        self.keys = keys
        self.values = catagories
        self.rows = self.getRows()
        self.cols = self.getColumns()
        
    def getRows(self) -> int:
        length = len(self.keys)
        return length


    def getColumns(self) -> int:
        length = len(self.values)
        return length

    def create(self, saveLocation):
        createFile = oxl.Workbook()
        activeFile = createFile.active
        activeFile.title = self.name

        i = 0
        while i < self.rows:
            activeFile[f"A{i+2}"]  =  self.keys[i]
            i += 1

        i = 0
        while i < self.cols:
            activeFile[f"{simple_i2a(i+1)}1"] = self.values[i][0]
            i += 1

        createFile.save(f"{saveLocation}/{self.name}.xlsx")

        



