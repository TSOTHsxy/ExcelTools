# -*- coding: utf-8 -*-

from script.utils.sequence import remove, merge, extract
from script.utils.sequence import mixed, replace, match

from script.utils.decorator import deploy


class SheetWrapper(object):
    """
    Worksheet wrapper.
    Extend some interfaces for read and write.
    """

    def __init__(self, worksheet):
        self.worksheet = worksheet
        self._init_all_attribute()

    @deploy('conf/guess.conf', ['字段关键字', '结尾关键字'])
    def _init_all_attribute(self, fields, ending):
        """
        Guess the attributes of the worksheet.
        Note: Guess attributes consumes a lot of system resources.
        """
        for row, cells in enumerate(self.worksheet.rows):
            values = replace(extract(cells), ' ', '')

            if mixed(fields, values):
                self._header, self._indexes = row, {
                    values[i]: i for i in range(len(values)) if values[i]
                }
                self._fields = [field for field, _ in sorted(
                    self._indexes.items(), key=lambda x: x[1]
                )]
                merge(remove(values, ''), fields)
                break
        else:
            raise AttributeError(
                'Fields attributes cannot be inferred.'
            )

        if (self.worksheet.nrows - 1) <= self._header:
            raise AttributeError('Ending attributes cannot be inferred.')

        for row in range(self.worksheet.nrows - 1, self._header, -1):
            if match(ending, extract(self.worksheet.row(row), True)):
                self._ending = row + 1
                break
        else:
            raise AttributeError(
                'Ending attributes cannot be inferred.'
            )

        return fields, ending

    def cell(self, row, field):
        """
        Returns the cell
        of the specified field in the specified row.
        """
        return self.worksheet.cell(row, self._indexes[field])

    def row(self, row, fields, mark=False):
        """
        Returns the cells
        of the specified fields in the specified row.
        """
        cells = [self.cell(row, field) for field in fields]
        return list(zip(fields, cells)) if mark else cells

    def rows(self, fields, mark=False):
        """
        Produces cells ​​for specified
        fields in worksheet by row (see: func: `row`).
        """
        for row in self.iter:
            yield self.row(row, fields, mark)

    @property
    def nrows(self) -> int:
        """
        Returns the count of rows in the worksheet.
        What is returned here is the number of rows of actual content rows.
        """
        return self._ending - self._header - 1

    @property
    def ncols(self) -> int:
        """
        Returns the count of cols in the worksheet.
        This returns the number of fields in the worksheet.
        """
        return len(self.fields)

    @property
    def title(self) -> str:
        """
        Returns the title of the current worksheet.
        A simple layer of proxy is added here.
        """
        return self.worksheet.title

    @title.setter
    def title(self, value):
        """
        Sets the title of the current worksheet.
        A simple layer of proxy is added here.
        """
        self.worksheet.title = value

    @property
    def fields(self):
        """Returns the fields in the worksheet."""
        return self._fields

    @property
    def iter(self):
        """Iterate over all valid line numbers."""
        return range(self._header + 1, self._ending)


class EmptyWrapper(object):
    """
    Empty worksheet wrapper.
    Extend some interfaces to adding records.
    """

    def __init__(self, worksheet):
        self.worksheet = worksheet
        self._indexes, self._nrows = {}, 2

    def add_row(self, cells, newline=True):
        """
        Write values ​​to the worksheet.
        The newline parameter specifies whether to wrap a line.
        """
        for field, cell in cells: self.worksheet.cell(
            self._nrows, self._indexes[field],
            cell.value
        )
        if newline: self._nrows += 1
        return self

    def add_rows(self, rows):
        """
        Write multiple rows of values ​​to tne worksheet.
        Wrap by default after writing a line.
        """
        for cells in rows: self.add_row(cells)
        return self

    def add_fields(self, fields):
        """
        Write fields to the worksheet.
        The fields are written to the second line in order,
        and duplicate fields are ignored.
        """
        for field in fields:
            if field in self._indexes:  continue
            index = len(self._indexes)

            self.worksheet.cell(1, index, field)
            self._indexes[field] = index

        return self

    def vary(self):
        """Conversion worksheet outer wrapping."""
        return SheetWrapper(self.worksheet)
