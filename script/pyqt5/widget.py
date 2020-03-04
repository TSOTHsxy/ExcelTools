# -*- coding: utf-8 -*-

from os.path import basename
from operator import eq

from PyQt5.QtWidgets import QTableWidget, QHeaderView, QTableWidgetItem, QMenu
from PyQt5.QtWidgets import QTreeWidget, QTreeWidgetItem
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor

from script.pyqt5.dialog import critical, input
from script.pyqt5.design import Dialog, hor_size, ver_size, col_size

from script.utils.sequence import eliminate
from script.utils.logger import debug
from script.excel.handler import iworkbook


class CellItem(QTableWidgetItem):
    def __init__(self, cell):
        super().__init__(str(cell.value))
        self.cell = cell
        self.alignment()

    def replace(self, old, new):
        value = str(self.cell.value).replace(old, new)
        super().setText(value)
        self.cell.value = value

    def alignment(self):
        self.setTextAlignment(Qt.AlignCenter)
        return self

    def protrude(self):
        self.setForeground(QBrush(QColor(255, 0, 0)))
        return self

    def disabled(self):
        self.setFlags(Qt.ItemIsEnabled)
        return self


class CellTable(QTableWidget):
    def __init__(self, parent, titles, row, col):
        super().__init__(parent)

        self.setFocusPolicy(Qt.NoFocus)
        self.itemChanged.connect(self._item_changed_listener)

        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._produce_right_menu)

        self.resize(hor_size, ver_size)
        self.setRowCount(row)
        self.setColumnCount(col)

        self.setHorizontalHeaderLabels(titles)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setDefaultAlignment(Qt.AlignCenter)

    def _produce_right_menu(self, pos):
        menu = QMenu()
        with open('res/qss/menu.qss') as file:
            menu.setStyleSheet(file.read())

        action = menu.addAction('替换文本')
        if menu.exec_(self.mapToGlobal(pos)) == action: self._replace_items()

    def _replace_items(self):
        content, proceed = input(self, '替换规则', '例：王>李；;')
        if not proceed or not content: return

        try:
            rules = [item.split('>') for item in content.split(';')]
            for item in self.selectedItems():
                for old, new in rules: item.replace(old, new)
        except:
            critical(self, '替换内容失败')
            debug('替换内容失败')

    @staticmethod
    def _item_changed_listener(item):
        if not isinstance(item, CellItem): return
        item.cell.value = item.text()


class SpecialTable(CellTable):
    def __init__(self, parent, titles, stuff, chooses, row, col):
        super().__init__(parent, titles, row, col)

        for row, tuple in enumerate(stuff):
            refer_cell, comp_cells, found_cells = tuple
            self.setItem(row, 0, CellItem(refer_cell).disabled())

            col = 1
            for cell in comp_cells:
                self.setItem(row, col, CellItem(cell))
                col += 1

            if not found_cells: continue
            for worksheet, _ in chooses:
                for comp, cell in zip(comp_cells, found_cells[worksheet]):
                    self.compare(comp, cell, row, col)
                    col += 1

    def compare(self, comp, cell, row, col):
        self.setItem(
            row, col, CellItem(cell) if eq(comp, cell)
            else CellItem(cell).protrude()
        )


class _HashableItem(QTreeWidgetItem):
    def __init__(self, parent, content):
        super().__init__(parent)
        self.setText(0, content)

    def __hash__(self):
        return hash(self.__str__())


class _CheckableItem(_HashableItem):
    def __init__(self, parent, content, checked=False):
        super().__init__(parent, content)
        self.setCheckState(0, Qt.Checked if checked else Qt.Unchecked)


class _WorksheetItem(_CheckableItem):
    def __init__(self, parent, workbook, worksheet, path):
        super().__init__(parent, worksheet.title)
        self.workbook, self.worksheet, self.path = workbook, worksheet, path


class _Tree(QTreeWidget):
    def __init__(self, labels, handler=None):
        super().__init__()
        self._handler = handler
        self.checkeds = []

        self.itemChanged.connect(self._item_changed_listener)
        self.itemDoubleClicked.connect(self._double_clicked_listener)

        self.setColumnWidth(0, col_size)
        self.setColumnCount(len(labels))

        self.setHeaderLabels(labels)
        self.setFocusPolicy(Qt.NoFocus)

    def clear(self):
        for item in list(self.checkeds): item.setCheckState(0, Qt.Unchecked)

    def _item_changed_listener(self, item):
        checked = item.checkState(0) == Qt.Checked
        isblank = item.text(1) == ''

        if not checked and not isblank:
            if self._handler: self._handler.__delete__(item)
            self.checkeds.remove(item)
            item.setText(1, '')

            for index, _item in enumerate(self.checkeds):
                _item.setText(1, str(index + 1))

        elif checked and isblank:
            if self._handler: self._handler.__append__(item)
            self.checkeds.append(item)
            item.setText(1, str(len(self.checkeds)))

    def _double_clicked_listener(self, item):
        pass


class WorkbookTree(_Tree):
    def __init__(self, labels, handler=None):
        super().__init__(labels, handler)

        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._generate_right_menu)

        self._load_tree = _HashableItem(self, '已加载')
        self._made_tree = _HashableItem(self, '已生成')
        self.expandAll()

    def load(self, path):
        try:
            workbook = iworkbook(path)
            parent = _HashableItem(self._load_tree, basename(path))
            for worksheet in workbook: _WorksheetItem(parent, workbook, worksheet, path)
        except:
            critical(self, '加载表格失败')
            debug('加载表格失败')

    def made(self, workbook):
        _WorksheetItem(self._made_tree, workbook, workbook.active, '')

    def _double_clicked_listener(self, item):
        if not isinstance(item, _WorksheetItem): return
        worksheet = item.worksheet
        dialog = Dialog(item.text(0))

        table = CellTable(dialog, worksheet.fields, worksheet.nrows, worksheet.ncols)
        for row, cells in enumerate(item.worksheet.rows(worksheet.fields)):
            for col, cell in enumerate(cells):
                table.setItem(row, col, CellItem(cell))

        dialog.exec_()

    def _generate_right_menu(self, pos):
        item = self.currentItem()
        if not isinstance(item, _WorksheetItem): return

        menu = QMenu()
        with open('res/qss/menu.qss') as file:
            menu.setStyleSheet(file.read())

        action = menu.addAction('修改表名')
        if menu.exec_(self.mapToGlobal(pos)) == action:
            content, proceed = input(self, '更改表名', '请输入更改后的表名')
            if not proceed or not content: return

            item.worksheet.title = content
            item.setText(0, content)


class FieldTree(_Tree):
    def __init__(self, labels, handler=None):
        super().__init__(labels, handler)
        self._public_tree = _HashableItem(self, '公共字段')
        self._privat_tree = _HashableItem(self, '私有字段')
        self._public, self._privat = {}, {}
        self.expandAll()

    def chooses(self, items):
        return [(item, eliminate([
            checked.text(0) for checked in self.checkeds
            if checked.parent() is self._public_tree
               or checked.parent() is self._privat[item]
        ])) for item in items]

    def __append__(self, item):
        self.clear()

        added = set(item.worksheet.fields)
        exist = set(self._public.keys())

        parent = _HashableItem(self._privat_tree, item.text(0))
        for field in added: _CheckableItem(parent, field)
        self._privat[item] = parent

        if exist and exist in added:  return
        if not exist: self.add_public(added)

        for field in exist - (exist & added):
            self._public_tree.removeChild(self._public[field])
            del self._public[field]

    def __delete__(self, item):
        self.clear()

        self._privat_tree.removeChild(self._privat[item])
        del self._privat[item]

        for _item in self._public.values():
            self._public_tree.removeChild(_item)
        self._public.clear()

        public = None
        for fields in [set(_item.worksheet.fields) for _item in self._privat]:
            public = public & fields if public else fields
        if public: self.add_public(public)

    def add_public(self, fields):
        for field in fields: self._public[field] = \
            _CheckableItem(self._public_tree, field)
