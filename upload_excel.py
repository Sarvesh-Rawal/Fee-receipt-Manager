from PyQt5.QtWidgets import QFileDialog
import pandas as pd


def upload_file(parent):
    file_dialog = QFileDialog()
    file_path, _ = file_dialog.getOpenFileName(parent, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
    return file_path
