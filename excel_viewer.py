import pandas as pd
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QMessageBox, QCheckBox, QPushButton
from PyQt5.QtCore import Qt

def display_excel_data(file_path, table_widget, selection_handler, print_handler):
    """
    Reads an Excel file, displays its data in a QTableWidget,
    and connects row checkboxes to a handler.

    Args:
        file_path (str): The path to the Excel file.
        table_widget (QTableWidget): The table widget to display the data.
        selection_handler (callable): A function to call when a checkbox state changes.
                                      It receives (state, row_index).
        print_handler (callable): A function to call when a row's print button is clicked.
                                  It receives (row_index).
    Returns:
        pd.DataFrame: The loaded DataFrame, or an empty DataFrame on error.
    """
    # Clear the table before loading new data to prevent old data from persisting
    table_widget.setRowCount(0)
    table_widget.setColumnCount(0)

    try:
        df = pd.read_excel(file_path)

        # Set column count to be DataFrame columns + checkbox + print button
        table_widget.setColumnCount(df.shape[1] + 2)
        table_widget.setRowCount(df.shape[0])

        # Set headers, including one for the checkbox column
        headers = ["Select"] + list(df.columns) + ["Action"]
        table_widget.setHorizontalHeaderLabels(headers)

        for row_idx in range(df.shape[0]):
            # --- Checkbox Column ---
            checkbox = QCheckBox()
            # Connect the checkbox signal to the handler provided by MainWindow
            checkbox.stateChanged.connect(lambda state, r=row_idx: selection_handler(state, r))
            table_widget.setCellWidget(row_idx, 0, checkbox)

            # --- Data Columns ---
            for col_idx in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iloc[row_idx, col_idx]))
                # Data is placed in column `col + 1` to make space for the checkbox
                table_widget.setItem(row_idx, col_idx + 1, item)

            # --- Print Button Column ---
            print_button = QPushButton("Print")
            print_button.clicked.connect(lambda checked, r=row_idx: print_handler(r))
            # Place button in the last column
            table_widget.setCellWidget(row_idx, df.shape[1] + 1, print_button)

        # Resize columns to fit content for better viewing
        table_widget.resizeColumnsToContents()
        return df  # Return the dataframe for state management
    except FileNotFoundError:
        QMessageBox.critical(None, "Error", f"File not found:\n{file_path}")
    except Exception as e:
        QMessageBox.critical(None, "Error", f"Could not load data from Excel file:\n{e}")

    return pd.DataFrame() # Return empty on failure