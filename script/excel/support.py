# -*- coding: utf-8 -*-

from shutil import copyfile
from uuid import uuid4

from xlutils.copy import copy as link
from xlrd import open_workbook


class XlsWorkbook(object):
    """
    Xls workbook.
    Establish a synchronization mechanism
    between read-only and write-only workbooks.
    """

    def __init__(self, path):
        self._path = 'cache/{}.xls'.format(uuid4().hex)
        copyfile(path, self._path)

        self.readable = open_workbook(self._path)
        self.writable = link(self.readable)

    def save(self, path):
        """
        Save workbook to disk.
        This is achieved by directly copying local cache.
        """
        # self.writable.save(path)
        copyfile(self._path, path)

    def flush(self):
        """Synchronize change to a read-only workbook."""
        self.writable.save(self._path)
        self.readable = open_workbook(self._path)

    def wsheet(self, index):
        """
        Returns a write-only
        worksheet based on the specified worksheet index.
        """
        return self.writable.get_sheet(index)

    def rsheet(self, index):
        """
        Returns a read-only
        worksheet based on the specified worksheet index.
        """
        return self.readable.sheet_by_index(index)
