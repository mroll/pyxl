import xlrd


def colmap(sheet):
    idxs  = range(sheet.ncols)
    names = map(str, sheet.row_values(0))
    
    return { pair[1]: pair[0] for pair in zip(idxs, names) }

def sheets(wb):
    return [Sheet(wb.sheet_by_index(i)) for i in range(wb.nsheets)]

class Sheet:
    def __init__(self, sheet):
        self.current  = 1
        self.sheet    = sheet
        self.colmap   = colmap(self.sheet)
        self.hdr      = self.sheet.row_values(0)

    def __getitem__(self, row):
        if type(row) == str:
            return self.sheet.col_values(self.colmap[row], start_rowx=1)
        if type(row) == tuple:
            row, col = row
            return self.sheet.cell(row, col).value
        else:
            return self.sheet.row_values(row)

    def __iter__(self):
        return self

    def next(self):
        if self.current >= self.sheet.nrows:
            current = 1
            raise StopIteration
        else:
            data = self.sheet.row_values(self.current)
            self.current += 1

            return data

    def __len__(self):
        return self.sheet.nrows

class Pyxl:
    def __init__(self, fname):
        self.wb     = xlrd.open_workbook(fname)
        self.sheets = sheets(self.wb)

    def __getitem__(self, idx):
        return self.sheets[idx]

    def __iter__(self):
        return self.sheets.__iter__()

    def __len__(self):
        return len(self.sheets)
