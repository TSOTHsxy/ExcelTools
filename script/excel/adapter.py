# -*- coding: utf-8 -*-

from abc import ABC, abstractmethod
from operator import eq

from script.utils.sequence import extract
from script.utils.filter import retype


class _CellAdapter(ABC):
    """
    Cell adapter.
    Provides unified interface for cell in worksheet.
    """

    def __hash__(self):
        return hash(self.__str__())

    def __eq__(self, other):
        if not isinstance(other, self.__class__): return False
        return eq(*retype(self.value, other.value))

    @property
    @abstractmethod
    def value(self):
        """Get the value held in the cell."""
        pass

    @value.setter
    @abstractmethod
    def value(self, value):
        """Set the value held in the cell."""
        pass


class _XlsCellAdapter(_CellAdapter):
    """
    Xls cell adapter.
    The bottom layer depends on the xlrd, xlwt modules.
    """

    def __init__(self, worksheet, row, col):
        self._worksheet = worksheet
        self._row, self._col = row, col

    @property
    def value(self):
        cell = self._worksheet.readable.cell(self._row, self._col)
        return cell.value if cell.value else ''

    @value.setter
    def value(self, value):
        self._worksheet.writable.write(self._row, self._col, value)
        self._worksheet.flush()


class _XlsxCellAdapter(_CellAdapter):
    """
    Xlsx cell adapter.
    The bottom layer depends on the openpyxl module.
    """

    def __init__(self, cell):
        self._cell = cell

    @property
    def value(self):
        return self._cell.value if self._cell.value else ''

    @value.setter
    def value(self, value):
        self._cell.value = value


class _SheetAdapter(ABC):
    """
    Worksheet adapter.
    Provides unified interface for worksheet.
    """

    def __init__(self):
        self._cell_cache = {}

    def cell(self, row, col, value=None):
        """
        Returns cells based on the specified cell index.
        First look in the cache for available cell.
        """
        if 0 > row >= self.nrows or 0 > col >= self.ncols:
            raise IndexError('Illegal cell index.')

        coordinate = (row, col)
        if coordinate not in self._cell_cache:
            self._cell_cache[coordinate] = self._adapt(row, col)

        cell = self._cell_cache[coordinate]
        if value is not None: cell.value = value
        return cell

    def row(self, row):
        """
        Returns the cells of the specified row.
        Maybe it is better to use a generator here.
        """
        return [
            self.cell(row, col)
            for col in range(self.ncols)
        ]

    @property
    def rows(self):
        """
        Produces all cells
        in the worksheet by row (see: func: `row`).
        """
        for row in range(self.nrows): yield self.row(row)

    def empty(self) -> bool:
        """
        Returns whether the worksheet is empty.
        Mainly for special wrapping of empty worksheets.
        """
        return not (self.nrows and self.ncols)

    @property
    @abstractmethod
    def nrows(self) -> int:
        """Returns the count of rows in the worksheet."""
        pass

    @property
    @abstractmethod
    def ncols(self) -> int:
        """Returns the count of cols in the worksheet."""
        pass

    @property
    @abstractmethod
    def title(self) -> str:
        """Returns the title of the current worksheet."""
        pass

    @title.setter
    @abstractmethod
    def title(self, value):
        """Sets the title of the current worksheet."""
        pass

    @abstractmethod
    def _adapt(self, row, col) -> _CellAdapter:
        pass


class XlsSheetAdapter(_SheetAdapter):
    """
    Xls worksheet adapter.
    The bottom layer depends on the xlrd, xlwt modules.
    """

    def __init__(self, workbook, index):
        super().__init__()
        self._workbook, self._index = workbook, index
        self.readable = self._workbook.rsheet(self._index)
        self.writable = self._workbook.wsheet(self._index)

    @property
    def nrows(self):
        return self.readable.nrows

    @property
    def ncols(self):
        return self.readable.ncols

    @property
    def title(self):
        return self.readable.name

    @title.setter
    def title(self, value):
        self.readable.name = value

    def flush(self):
        """Synchronize change to a read-only worksheet."""
        self._workbook.flush()
        self.readable = self._workbook.rsheet(self._index)

    def _adapt(self, row, col):
        return _XlsCellAdapter(self, row, col)


class XlsxSheetAdapter(_SheetAdapter):
    """
    Xlsx worksheet adapter.
    The bottom layer depends on the openpyxl module.
    """

    def __init__(self, worksheet):
        super().__init__()
        self._worksheet = worksheet

    @property
    def nrows(self):
        return self._worksheet.max_row

    @property
    def ncols(self):
        return self._worksheet.max_column

    @property
    def title(self):
        return self._worksheet.title

    @title.setter
    def title(self, value):
        self._worksheet.title = value

    def empty(self):
        if not super().empty():
            for row in self.rows:
                if extract(row, True): return False
        return True

    def _adapt(self, row, col):
        # In openpyxl module, cell index starts from 1.
        return _XlsxCellAdapter(self._worksheet.cell(row + 1, col + 1))


class BookAdapter(object):
    """
    Workbook adapter.
    Provide unified interface for workbook.
    """

    def __init__(self, workbook, allsheet):
        self._workbook = workbook
        self._allsheet = allsheet

    @property
    def active(self):
        """Returns the active worksheet."""
        return self._allsheet[0]

    def save(self, path):
        """
        Save the workbook to
        the specified path on the hard disk.
        """
        self._workbook.save(path)

    def vary(self):
        """
        Conversion worksheet outer wrapping.
        To adapt to the special needs of this project.
        """
        for index, worksheet in enumerate(self._allsheet):
            self._allsheet[index] = worksheet.vary()

    def __iter__(self):
        return iter(self._allsheet)

    def __len__(self):
        return len(self._allsheet)

    def __getitem__(self, key):
        return self._allsheet[key]
