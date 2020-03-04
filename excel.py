# -*- coding: utf-8 -*-

from functools import wraps
from re import split, findall
from sys import argv, exit

from shutil import rmtree as rmdir
from os import mkdir

from PyQt5.QtWidgets import QApplication, QGridLayout

from script.utils.filter import filter
from script.pyqt5.widget import *
from script.pyqt5.design import *
from script.pyqt5.dialog import *


def _create():
    workbook = iworkbook()
    return workbook, workbook.active


def _catch(func):
    @wraps(func)
    def proxy(self):
        try:
            return func(self)
        except:
            critical(self, '发生未知错误')
            debug('发生未知错误')

    return proxy


class ExcelTools(Window):
    def __init__(self):
        super().__init__('ExcelTool')
        self.setLayout(QGridLayout())
        self.loaded = set()

        self.field_tree = FieldTree(['字段', '编号'])
        self.sheet_tree = WorkbookTree(['表格', '编号'], self.field_tree)

        self.layout().addWidget(
            Vertical(self.sheet_tree, Button(self, '加载表格', self._load_excels))
            , 0, 0)

        self.layout().addWidget(
            Vertical(self.field_tree, Button(self, '重置选择', self._clear_check))
            , 0, 1)

        actions = Vertical(
            Button(self, '提取多行', self._extract_rows),
            Button(self, '提取多列', self._extract_cols),
            Button(self, '比较数据', self._compare_cell),
            Button(self, '筛选数据', self._screen_cells),
            Button(self, '保存操作', self._save_options)
        )

        actions.layout().addStretch()
        self.layout().addWidget(actions, 0, 2)

    def closeEvent(self, QCloseEvent):
        rmdir('cache/')
        mkdir('cache/')

    @_catch
    def _load_excels(self):
        types = 'Excel Files (*.xls *.xlsx)'
        files = set(select_files(self, types)) - self.loaded
        for filename in files: self.sheet_tree.load(filename)
        self.loaded.update(files)

    @_catch
    def _clear_check(self):
        self.sheet_tree.clear()
        self.field_tree.clear()

    @_catch
    def _make_chooses(self):
        if not self.sheet_tree.checkeds or not self.field_tree.checkeds:
            critical(self, '请指定需要处理的表格和字段')
            return None
        return [
            (item.worksheet, fields) for item, fields in
            self.field_tree.chooses(self.sheet_tree.checkeds)
            if fields
        ]

    @_catch
    def _extract_rows(self):
        chooses = self._make_chooses()
        if not chooses: return
        workbook, cursheet = _create()

        for worksheet, fields in chooses: cursheet.add_fields(fields).add_rows([
            row for row in worksheet.rows(fields, True)
        ])

        workbook.vary()
        self.sheet_tree.made(workbook)

    @_catch
    def _extract_cols(self):
        chooses = self._make_chooses()
        if not chooses: return

        target, fields = chooses[0][0], chooses[0][1]
        del chooses[0]

        public = set(fields)
        for _, _fields in chooses:
            public = public & set(_fields)

        if len(public) == 0:
            critical(self, '请指定基准字段')
            return

        workbook = iworkbook()
        cursheet = workbook.active.add_fields(fields)

        for index, tuple in enumerate(list(chooses)):
            private = [field for field in tuple[1] if field not in public]
            cursheet.add_fields(private)
            chooses[index] = (tuple[0], private)

        for row in target.iter:
            cursheet.add_row(target.row(row, fields, True), False)
            contrast = target.row(row, public)

            cursheet.add_rows([
                worksheet.row(_row, _fields, True)
                for worksheet, _fields in chooses for _row in worksheet.iter
                if eq(contrast, worksheet.row(_row, public))
            ])

        workbook.vary()
        self.sheet_tree.made(workbook)

    @_catch
    def _compare_cell(self):
        chooses = self._make_chooses()
        if not chooses: return

        if len(chooses) == 1:
            critical(self, '请指定两张表格')
            return

        target, fields = chooses[0][0], chooses[0][1]
        public, refer = set(chooses[0][1]), fields[0]
        del chooses[0]

        for _, _fields in chooses:
            public = public & set(_fields)

        if len(public) == 0:
            critical(self, '请指定公共字段')
            return

        if refer not in public:
            critical(self, '请指定基准字段')
            return

        public = sorted(public, key=fields.index)
        public.remove(refer)

        stuff = []

        for row in target.iter:
            refer_cell = target.cell(row, refer)

            stuff.append((refer_cell, target.row(row, public), {
                worksheet: worksheet.row(_row, public)
                for worksheet, _ in chooses for _row in worksheet.iter
                if eq(refer_cell, worksheet.cell(_row, refer))
            }))

        if not stuff:
            critical(self, '当前无差异字段')
            return

        titles = ['基准字段'] + ['基准表.' + field for field in public]
        for worksheet, fields in chooses: titles.extend([
            '{}.{}'.format(worksheet.title, field)
            for field in public
        ])

        nrows = len(stuff)
        ncols = (len(chooses) + 1) * (len(public)) + 1

        dialog = Dialog('比较结果')
        SpecialTable(dialog, titles, stuff, chooses, nrows, ncols)
        dialog.exec_()

    @_catch
    def _screen_cells(self):
        chooses = self._make_chooses()
        if not chooses: return
        workbook, cursheet = _create()

        content, proceed = input(self, '筛选规则', '例：金额>50;')
        if not proceed: return

        if content == '':
            critical(self, '请输入筛选规则')
            return

        regex, rules = '[>|>=|<|<=|=|!=]', {}
        for item in content.split(';'):
            symb, expr = findall(regex, item)[0], split(regex, item)
            if len(expr) != 2: continue

            target = expr[0]
            if target not in rules: rules[target] = []
            rules[target].append((symb, expr[1]))

        rows = [
            (fields, row) for worksheet, fields in chooses
            for row in worksheet.rows(fields, True)
            if filter(row, rules)
        ]

        if not rows:
            critical(self, '筛选结果为空值')
            return

        for fields, row in rows:
            cursheet.add_fields(fields).add_row(row)

        workbook.vary()
        self.sheet_tree.made(workbook)

    @_catch
    def _save_options(self):
        for item in self.sheet_tree.checkeds:
            if not item.path: item.path = get_filepath(self)
            item.workbook.save(item.path)


if __name__ == '__main__':
    app = QApplication(argv)
    try:
        frame = ExcelTools()
        frame.show()
    except:
        debug('加载页面失败')
    exit(app.exec_())
