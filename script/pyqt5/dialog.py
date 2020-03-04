# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import QFileDialog, QMessageBox, QInputDialog


def input(env, title, prompt):
    return QInputDialog.getText(env, title, prompt)


def get_filepath(env, path='./'):
    return QFileDialog.getSaveFileName(env, '保存文件', path)[0]


def select_files(env, path='./', filter='All files(*.*)'):
    return QFileDialog.getOpenFileNames(env, '选择文件', path, filter)[0]


def critical(env, prompt):
    QMessageBox.critical(env, '错误', prompt, QMessageBox.Yes)
