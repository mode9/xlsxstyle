# This Python file uses the following encoding: utf-8
import datetime
import os

from PyQt5 import QtCore
from PySide2.QtWidgets import QTextBrowser, QProgressBar
from openpyxl import Workbook

from handlers import WorkSheetHandler


class CreateThread(QtCore.QThread):
    finished = QtCore.pyqtSignal()

    def __init__(self, org_wb: Workbook, target_wb: Workbook, text_browser: QTextBrowser, parent=None):
        super().__init__(parent=parent)
        self.org_wb = org_wb
        self.target_wb = target_wb
        self.text_browser = text_browser

    def get_handler(self, workbook: Workbook):
        SHEET_NAME = "GI"
        hdlr = WorkSheetHandler(workbook, sheet_name=SHEET_NAME)
        self.text_browser.insertPlainText(f"INFO: 시트 로딩 완료 ({SHEET_NAME})\n")
        return hdlr

    def get_new_filename(self) -> str:
        today: datetime.date = datetime.date.today()
        DIR: str = os.path.abspath(os.path.curdir)
        file_dir: str = os.path.join(DIR, today.strftime("%Y%m%d"))
        file_dir: str = self.check_unique_filename(file_dir)
        return file_dir

    def check_unique_filename(self, filename: str, extra: str = '') -> str:
        if os.path.isfile(filename + extra):
            extra: str = str(int(extra) + 1) if extra else '1'
            return self.check_unique_filename(filename, extra)
        return filename

    def run(self) -> None:
        org_handler = self.get_handler(self.org_wb)
        target_handler = self.get_handler(self.target_wb)
        org_handler.copy_styles(target_handler)
        self.text_browser.insertPlainText("INFO: 스타일 복사 완료\n")
        fn = self.get_new_filename() + '.xlsx'
        self.target_wb.save(fn)
        self.text_browser.insertPlainText(f"INFO: 파일 생성 완료 ({fn}) \n")
        self.finished.emit()
        self.org_wb.close()
        self.target_wb.close()


class ProgressThread(QtCore.QThread):
    def __init__(self, pg_bar: QProgressBar, parent=None):
        super().__init__(parent=parent)
        self.pg_bar = pg_bar
        self._status = False

    def run(self):
        self.pg_bar.setRange(0, 100)

    def toggle_status(self):
        self._status = not self._status
        maximum = 0 if self._status else 100
        self.pg_bar.setRange(0, maximum)

    @property
    def status(self):
        return self._status

    def __del__(self):
        self.wait()