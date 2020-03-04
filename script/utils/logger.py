# -*- coding: utf-8 -*-

from logging.handlers import TimedRotatingFileHandler
from os import mknod
from os.path import exists

from time import strftime

import logging

_path = 'logs/{}.log'.format(strftime("%Y%m%d"))
if not exists(_path): mknod(_path)

logger = logging.getLogger('ExcelTools')
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter(
    '%(asctime)s - %(name)s[line:%(lineno)d]' +
    ' - %(levelname)s: %(message)s'
)

_console_handler = logging.StreamHandler()
_console_handler.setLevel(logging.DEBUG)
_console_handler.setFormatter(formatter)
logger.addHandler(_console_handler)

_file_handler = TimedRotatingFileHandler(filename=_path)
_file_handler.setLevel(logging.DEBUG)
_file_handler.setFormatter(formatter)
logger.addHandler(_file_handler)


def debug(logs):
    logger.debug(logs, exc_info=True)
    logger.debug('\n')
