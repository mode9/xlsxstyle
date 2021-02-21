# This Python file uses the following encoding: utf-8
from PyQt5 import QtCore

from helpers import get_workbook


class Worker(QtCore.QObject):
    finished = QtCore.pyqtSignal()
    workbook = QtCore.pyqtSignal(object)
    failed = QtCore.pyqtSignal()

    def __init__(self, filename, text_browser, parent=None):
        super().__init__(parent=parent)
        self.filename = filename
        self.text_browser = text_browser

    def run(self):
        try:
            wb = get_workbook(self.filename, self.text_browser)
        except Exception:
            self.failed.emit()
            self.finished.emit()
            return

        self.workbook.emit(wb)
        self.finished.emit()