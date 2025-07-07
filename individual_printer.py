import os
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QWidget
from pdf_generator import create_receipt_pdf

def print_single_receipt_from_df(parent: QWidget, df: pd.DataFrame, row_index: int, student_name_column: str, print_file_handler: callable):
    """
    Handles the logic for generating, saving, and printing a single receipt.

    Args:
        parent (QWidget): The parent widget for dialogs.
        df (pd.DataFrame): The DataFrame containing all the data.
        row_index (int): The integer index of the row in the DataFrame to print.
        student_name_column (str): The name of the column containing student names.
        print_file_handler (callable): A function to call to print the generated PDF file.
                                       It receives the full file path.
    """
    if df is None or df.empty:
        QMessageBox.warning(parent, "No Data", "No data loaded to print from.")
        return

    save_dir = QFileDialog.getExistingDirectory(parent, "Select Directory to Save Receipt")
    if not save_dir:
        return

    # Use .iloc to get the row by its integer position from the DataFrame
    row_data = df.iloc[row_index]
    # The original DataFrame index is needed for a consistent filename
    original_df_index = row_data.name

    # Create filename
    if student_name_column in row_data and pd.notna(row_data[student_name_column]):
        student_name = str(row_data[student_name_column])
        sanitized_name = "".join(c for c in student_name if c.isalnum() or c in (' ', '_')).rstrip()
        file_name = f"receipt_{sanitized_name}_{original_df_index}.pdf"
    else:
        file_name = f"receipt_{original_df_index}.pdf"

    full_path = os.path.join(save_dir, file_name)

    if create_receipt_pdf(row_data, full_path):
        print_file_handler(full_path)
        QMessageBox.information(parent, "Success", f"Successfully saved receipt:\n{os.path.basename(full_path)}")
    else:
        QMessageBox.warning(parent, "PDF Error", "Failed to create the PDF receipt.")