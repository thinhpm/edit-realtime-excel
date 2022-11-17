import xlwings as xw


class Documents:
    def __init__(self):
        pass


class Excel(Documents):
    def __init__(self, file_location=""):
        super().__init__()
        self.file_location = file_location
        self.wb = xw.Book(file_location)
        self.sheet = self.wb.sheets[0]
        self.cell = None

    def cell_active(self):
        cell_range = self.wb.app.selection
        row_num = cell_range.row
        col_num = cell_range.column
        print(f"Cell current: {col_num} - {row_num}")
        self.cell = cell_range

    def cell_left(self):
        next_range_col = self.cell.column
        next_range_row = self.cell.row - 1

        self.cell = self.sheet.range((next_range_col, next_range_row))

    def cell_right(self):
        next_range_col = self.cell.column
        next_range_row = self.cell.row + 1

        self.cell = self.sheet.range((next_range_col, next_range_row))

    def cell_up(self):
        next_range_col = self.cell.column - 1
        next_range_row = self.cell.row

        self.cell = self.sheet.range((next_range_col, next_range_row))

    def cell_down(self):
        next_range_col = self.cell.column + 1
        next_range_row = self.cell.row

        self.cell = self.sheet.range((next_range_col, next_range_row))

    def cell_set(self, value):
        # self.sheet.range('A2').value = value
        self.cell.value = value