import time

import openpyxl


def get_workbook(filename, text_browser):
    start = time.time()
    wb = openpyxl.load_workbook(filename)
    runtime = get_working_time(start)
    text_browser.insertPlainText(
        f"INFO: 파일읽기 {filename.split('/')[-1]} (runtime: {runtime} sec)\n"
    )
    return wb


def get_working_time(start_time: float):
    curr_time = time.time()
    return round(curr_time - start_time, 2)
