# This Python file uses the following encoding: utf-8
import os
import sys
import time

from PyQt5 import QtCore
from PySide2.QtCore import QFile
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QApplication, QWidget, QFileDialog, \
    QPushButton, QLineEdit, QMessageBox, QTextBrowser, QProgressBar
from openpyxl import Workbook

from threads import CreateThread, ProgressThread
from workers import Worker


class XlsxStyler(QWidget):
    def __init__(self):
        super(XlsxStyler, self).__init__()
        self.load_ui()
        self.org_button: QPushButton = self.findChild(QPushButton, "org_button")
        self.target_button: QPushButton = self.findChild(QPushButton, "target_button")
        self.create_button: QPushButton = self.findChild(QPushButton, "create_button")
        self.org_edit: QLineEdit = self.findChild(QLineEdit, "org_edit")
        self.target_edit: QLineEdit = self.findChild(QLineEdit, "target_edit")
        self.text_browser: QTextBrowser = self.findChild(QTextBrowser, "textBrowser")
        self.pg_bar: QProgressBar = self.findChild(QProgressBar, "progressBar")
        self.pg_thread = ProgressThread(self.pg_bar)
        # self.pg_thread.change_value.connect(self._set_pg_range)
        self.pg_thread.start()
        self.org_button.clicked.connect(self.org_dialog)
        self.target_button.clicked.connect(self.target_dialog)
        self.create_button.clicked.connect(self.createExcel)
        self.org_name = ''
        self.target_name = ''
        self.org_wb = None
        self.target_wb = None

    def load_ui(self):
        loader = QUiLoader()
        path = os.path.join(os.path.dirname(__file__), "form.ui")
        ui_file = QFile(path)
        ui_file.open(QFile.ReadOnly)
        loader.load(ui_file, self)
        ui_file.close()

    def _set_edit_text(self, target: str, text: str) -> None:
        getattr(self, f'{target}_edit').setText(text)

    def _set_workbook(self, target: str, wb: Workbook) -> None:
        setattr(self, f'{target}_wb', wb)
        self.create_button.setDisabled(False)

    def set_org_wb(self, wb):
        self._set_workbook(target='org', wb=wb)

    def set_target_wb(self, wb):
        self._set_workbook(target='target', wb=wb)

    def _open_dialog(self, target: str):
        wb_callback = getattr(self, f"set_{target}_wb")
        name = QFileDialog.getOpenFileName(self, 'Open file', './')[0]
        if not name:
            return
        elif name.split(".")[-1] not in ("xlsx", "xlsm", "xlsb", "xls"):
            QMessageBox.critical(self, 'error', '엑셀 파일이 아닙니다.')
            return

        self.create_button.setDisabled(True)
        self.pg_thread.toggle_status()
        setattr(self, f'{target}_thread', QtCore.QThread())
        setattr(self, f'{target}_worker', Worker(name, self.text_browser))
        worker: Worker = getattr(self, f'{target}_worker')
        thread: QtCore.QThread = getattr(self, f'{target}_thread')
        worker.moveToThread(thread)

        thread.started.connect(worker.run)
        worker.finished.connect(thread.quit)
        worker.finished.connect(self.pg_thread.toggle_status)
        worker.failed.connect(lambda: QMessageBox.critical(self, 'error', '해당 파일이 존재하지 않거나 잘못된 파일입니다.'))
        worker.workbook.connect(wb_callback)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.start()
        self._set_edit_text(target, name)
        setattr(self, f"{target}_name", name)

    def org_dialog(self) -> None:
        self._open_dialog('org')

    def target_dialog(self) -> None:
        self._open_dialog('target')

    def createExcel(self) -> None:
        if not self.org_name or not self.target_name:
            QMessageBox.critical(self, 'error', '파일을 선택해주세요.')
            return
        self.create_button.setDisabled(True)
        self.text_browser.insertPlainText("INFO: 파일 생성 작업 시작\n")
        start = time.time()
        self.create_thread = CreateThread(self.org_wb, self.target_wb, self.text_browser)
        self.create_thread.started.connect(self.pg_thread.toggle_status)
        self.create_thread.finished.connect(lambda: self.insert_text(
            f"******** 총 작업시간: {round(time.time() - start, 2)}초 ********\n"
        ))
        self.create_thread.finished.connect(self.pg_thread.toggle_status)
        self.create_thread.finished.connect(lambda: self.create_button.setDisabled(False))
        self.create_thread.start()

    def insert_text(self, text: str):
        self.text_browser.insertPlainText(text)


if __name__ == "__main__":
    app = QApplication([sys.executable])
    widget = XlsxStyler()
    widget.show()
    sys.exit(app.exec_())
