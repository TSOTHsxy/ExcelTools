# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import QWidget, QDesktopWidget, QDialog
from PyQt5.QtWidgets import QVBoxLayout, QPushButton
from PyQt5.QtGui import QIcon

style, icon = 'res/qss/style.qss', 'res/img/icon.png'
hor_size, ver_size, col_size = 1000, 600, 300


def _init_basic_style(widget, title):
    widget.setWindowTitle(title)

    widget.setFixedSize(hor_size, ver_size)
    widget.resize(hor_size, ver_size)

    with open(style) as file: widget.setStyleSheet(file.read())
    widget.setWindowIcon(QIcon(icon))


class Vertical(QWidget):
    def __init__(self, *widgets):
        super().__init__()

        self.setLayout(QVBoxLayout())
        for widget in widgets: self.layout().addWidget(widget)


class Window(QWidget):
    def __init__(self, title):
        super().__init__()

        _init_basic_style(self, title)
        env, app = QDesktopWidget().screenGeometry(), self.geometry()
        self.move((env.width() - app.width()) / 2, (env.height() - app.height()) / 2)


class Dialog(QDialog):
    def __init__(self, title):
        super().__init__()
        _init_basic_style(self, title)


class Button(QPushButton):
    def __init__(self, parent, text, callback):
        super().__init__(text, parent)
        self.clicked.connect(callback)
