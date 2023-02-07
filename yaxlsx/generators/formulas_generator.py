from xlsxwriter.utility import xl_range_abs, xl_rowcol_to_cell_fast

from base import XlsxCell, XlsxCellsRange, XlsxFormula


class XlsxFormulasGenerator:
    def __init__(self):
        pass

    @staticmethod
    def if_(condition, then=None, else_=None) -> XlsxFormula:
        if then is None:
            then = 1
        if else_ is None:
            else_ = 0

        return XlsxFormula("IF", condition, then, else_)

    @staticmethod
    def concat(*values) -> XlsxFormula:
        return XlsxFormula("CONCATENATE", *values)

    @staticmethod
    def contents(*values) -> XlsxFormula:
        return XlsxFormula("CELL", "contents", *values)

    @staticmethod
    def sum(*values) -> XlsxFormula:
        return XlsxFormula("SUM", *values)

    @staticmethod
    def left(value, num_chars: int = 1) -> XlsxFormula:
        """
        Extracts a given number of characters from the left side of a supplied string value.
        If num_chars is greater than the number of characters available, function returns original string.

        :param value: original string or cell reference
        :param num_chars: the number of characters to extract
        """
        return XlsxFormula("LEFT", value, num_chars)

    @staticmethod
    def right(value, num_chars: int = 1) -> XlsxFormula:
        """
        Extracts a given number of characters from the right side of a supplied string value.
        If num_chars is greater than the number of characters available, function returns original string.

        :param value: original string or cell reference
        :param num_chars: the number of characters to extract
        """
        return XlsxFormula("RIGHT", value, num_chars)

    @staticmethod
    def find(substring, value) -> XlsxFormula:
        """
        Returns the index of one text string inside another.
        When the text is not found, it returns a #VALUE error.

        :param substring:
        :param value:
        :return:
        """
        return XlsxFormula("FIND", substring, value)

    @staticmethod
    def len(value) -> XlsxFormula:
        """
        Returns the length of a given text string (the number of characters).
        May also count characters in numbers, but number formatting is not included.

        :param value:
        :return:
        """
        return XlsxFormula("LEN", value)

    @staticmethod
    def isnumber(item) -> XlsxFormula:
        """
        Returns boolean result for condition: is cell/value contains a number

        You can use ISNUMBER to check that a cell contains a numeric value,
        or that the result of another function is a number.

        :param item:
        :return:
        """
        return XlsxFormula("ISNUMBER", item)

    @staticmethod
    def value(item) -> XlsxFormula:
        """
        Converts string that appears in a recognized format (number, date, time format) into a numeric value.

        :param item:
        :return:
        """
        return XlsxFormula("VALUE", item)

    @staticmethod
    def just(item):
        return XlsxFormula("", item)

    @staticmethod
    def sum_range(first_row: int, first_col: int, last_row: int, last_col: int) -> XlsxFormula:
        sum_range = XlsxCellsRange(
            xl_range_abs(first_row=first_row, first_col=first_col, last_row=last_row, last_col=last_col)
        )
        return XlsxFormula("SUM", sum_range)

    @staticmethod
    def sum_in_one_col(col: int, rows: list[int]) -> XlsxFormula:
        cells_list = [XlsxCell(xl_rowcol_to_cell_fast(col=col, row=row)) for row in rows]
        return XlsxFormula("SUM", *cells_list)

    @staticmethod
    def sum_in_one_row(row: int, cols: list[int]) -> XlsxFormula:
        cells_list = [XlsxCell(xl_rowcol_to_cell_fast(col=col, row=row)) for col in cols]
        return XlsxFormula("SUM", *cells_list)

    # RTR ------------------
    def __rtr_slice_value(self, cell, side, position) -> XlsxFormula:
        cell_left_value = self.value(side(cell, position))

        return self.if_(
            condition=self.isnumber(cell_left_value),
            then=cell_left_value,
            else_=0
        )

    def rtr_plan_value(self, cell):
        return self.__rtr_slice_value(cell, self.left, self.find(' / ', cell))

    def rtr_fact_value(self, cell):
        return self.__rtr_slice_value(cell, self.right, self.len(cell) - self.find(' / ', cell) - 1)

    def rtr_concat(self, plan_value, fact_value):
        return self.concat(plan_value, " / ", fact_value)

    def rtr_sum(self, *cells):
        plan_cells = [self.rtr_plan_value(cell) for cell in cells]
        fact_cells = [self.rtr_fact_value(cell) for cell in cells]

        return self.rtr_concat(self.sum(*plan_cells), self.sum(*fact_cells))

    def rtr_sum_in_one_col(self, col, rows: list[int]) -> XlsxFormula:
        cells = [XlsxCell(xl_rowcol_to_cell_fast(col=col, row=row)) for row in rows]
        return self.rtr_sum(*cells)

    def rtr_sum_in_one_row(self, row, cols: list[int]) -> XlsxFormula:
        cells = [XlsxCell(xl_rowcol_to_cell_fast(col=col, row=row)) for col in cols]
        return self.rtr_sum(*cells)
