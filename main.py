import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QTableWidget, QMessageBox,
                             QLineEdit, QGroupBox)
from PyQt5.QtCore import Qt
from upload_excel import upload_file  # Import the function
from excel_viewer import display_excel_data  # Import the new function
from pdf_generator import create_receipt_pdf # Import the new PDF generator
from table_filter import filter_table_by_name # Import the new filter function
from individual_printer import print_single_receipt_from_df # Import the new individual print logic


class MainWindow(QWidget):
    # --- Class Level Configuration ---
    # IMPORTANT: Change this value to match the exact column header for student names in your Excel file.
    STUDENT_NAME_COLUMN = 'Name'

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel File Viewer")
        self.setMinimumSize(1120, 600)

        # --- State Management ---
        self.name_column_table_index = -1  # The index of the name column in the QTableWidget
        self.df = None
        self.selected_rows = set()

        # --- Widgets ---
        self.upload_button = QPushButton("Upload Excel File")
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search by Name...")
        self.print_receipts_button = QPushButton("Print Receipt(s)")
        self.table_widget = QTableWidget()
        # Make the table read-only to prevent accidental edits
        self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)
        # Add padding to all data cells for better readability
        self.table_widget.setStyleSheet("QTableWidget::item { padding: 8px; }")

        # --- Layout ---
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.upload_button)

        # Use a QGroupBox for a visually and structurally robust container
        table_group_box = QGroupBox("Data Table")
        table_layout = QVBoxLayout(table_group_box)
        table_layout.addWidget(self.search_bar)
        table_layout.addWidget(self.table_widget, 1)

        # Add the group box to the main layout with a stretch factor.
        # This makes the entire table area expand and shrink with the window.
        main_layout.addWidget(table_group_box, 1)
        main_layout.addWidget(self.print_receipts_button)

        # --- Connections ---
        self.upload_button.clicked.connect(self.upload_file)
        self.print_receipts_button.clicked.connect(self.print_receipts)
        self.search_bar.textChanged.connect(self.on_search_text_changed)

    def upload_file(self):
        """Handles file upload and triggers table population."""
        # Clear previous state
        self.name_column_table_index = -1
        self.selected_rows.clear()
        self.df = None
        self.search_bar.clear()

        file_path = upload_file(self)
        if file_path:
            # Pass the selection handler method to the display function
            # and store the returned DataFrame
            self.df = display_excel_data(file_path, self.table_widget, self.on_selection_changed, self.print_single_receipt)
            if self.df is not None and not self.df.empty:
                if self.STUDENT_NAME_COLUMN in self.df.columns:
                    # Get the index from the DataFrame and add 1 for the "Select" column in the table
                    self.name_column_table_index = self.df.columns.get_loc(self.STUDENT_NAME_COLUMN) + 1
                else:
                    self.name_column_table_index = -1
                    QMessageBox.warning(self, "Column Not Found",
                                        f"The column '{self.STUDENT_NAME_COLUMN}' was not found.\n\n"
                                        "Search by name and PDF naming may not work as expected.")

    def on_selection_changed(self, state, row_index):
        """
        This method is called from excel_viewer when a checkbox is toggled.
        It manages the set of selected row indices.
        """
        if state == Qt.Checked:
            self.selected_rows.add(row_index)
        else:
            self.selected_rows.discard(row_index)

    def on_search_text_changed(self, text):
        """
        Called when the text in the search bar changes.
        Filters the table based on the new text.
        """
        # Only filter if the name column was successfully found on upload
        if self.name_column_table_index > 0:
            filter_table_by_name(self.table_widget, text, name_column_index=self.name_column_table_index)

    def _print_file(self, filepath):
        """
        Sends a file to the default printer.
        This implementation is primarily for Windows.
        """
        try:
            if sys.platform != "win32":
                QMessageBox.information(self, "Printing Not Supported",
                                      f"Automatic printing is not supported on your operating system ({sys.platform}).\n"
                                      f"You can find the saved PDF here:\n{filepath}")
                return

            os.startfile(filepath, "print")
        except OSError as e:
            if e.winerror == 1155:  # No application associated for printing
                reply = QMessageBox.question(self, "Printing Error",
                                             "Could not print automatically because no default PDF printing application is set up.\n\n"
                                             "Would you like to open the file to print it manually?",
                                             QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if reply == QMessageBox.Yes:
                    try:
                        os.startfile(filepath)
                    except Exception as open_e:
                        QMessageBox.warning(self, "File Open Error", f"Could not open the file.\nError: {open_e}")
            else:
                QMessageBox.warning(self, "Print Error", f"Could not print file: {os.path.basename(filepath)}\nError: {e}")
        except Exception as e:
            QMessageBox.warning(self, "Print Error", f"Could not print file: {os.path.basename(filepath)}\nError: {e}")

    def print_single_receipt(self, row_index):
        """Wrapper to call the individual receipt printing logic from the new file."""
        print_single_receipt_from_df(
            parent=self,
            df=self.df,
            row_index=row_index,
            student_name_column=self.STUDENT_NAME_COLUMN,
            print_file_handler=self._print_file
        )

    def print_receipts(self):
        """Converts selected row data into individual PDF receipts, saves, and prints them."""
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "No Data", "Please upload an Excel file first.")
            return

        if not self.selected_rows:
            QMessageBox.information(self, "No Selection", "Please select one or more rows using the checkboxes.")
            return

        # Ask user for a directory to save the files
        save_dir = QFileDialog.getExistingDirectory(self, "Select Directory to Save Receipts")

        if not save_dir:  # User cancelled the dialog
            return

        # Filter the selected rows to only include those currently visible
        visible_selected_rows = {
            row_idx for row_idx in self.selected_rows
            if not self.table_widget.isRowHidden(row_idx)
        }

        if not visible_selected_rows:
            QMessageBox.information(self, "No Visible Selection",
                                    "You have selected rows, but they are hidden by the current search filter.\n\n"
                                    "Please clear the search or change the filter to print receipts.")
            return

        selected_df = self.df.iloc[sorted(list(visible_selected_rows))]

        success_count = 0
        error_count = 0

        for index, row in selected_df.iterrows():
            # Create a filename from the student's name using the class constant
            if self.STUDENT_NAME_COLUMN in row and pd.notna(row[self.STUDENT_NAME_COLUMN]):
                student_name = str(row[self.STUDENT_NAME_COLUMN])
                # Sanitize the name to make it a valid filename by removing invalid characters.
                # This keeps letters, numbers, spaces, and underscores.
                sanitized_name = "".join(c for c in student_name if c.isalnum() or c in (' ', '_')).rstrip()
                # Add the original index to ensure the filename is unique.
                file_name = f"receipt_{sanitized_name}_{index}.pdf"
            else:
                # Fallback to the old naming scheme if the name column is missing or empty.
                file_name = f"receipt_{index}.pdf"

            full_path = os.path.join(save_dir, file_name)

            if create_receipt_pdf(row, full_path):
                success_count += 1
                self._print_file(full_path)
            else:
                error_count += 1

        # Provide feedback to the user
        summary_message = f"Successfully saved: {success_count} receipts."
        if error_count > 0:
            summary_message += f"\nFailed to save: {error_count} receipts. See console for errors."
        QMessageBox.information(self, "PDF Generation Complete", summary_message)

        # Unselect all checkboxes after the operation is complete.
        # We iterate over a copy of the set because unchecking the box will
        # trigger on_selection_changed, which modifies the set.
        for row_index in list(self.selected_rows):
            checkbox = self.table_widget.cellWidget(row_index, 0)
            if checkbox:
                checkbox.setChecked(False)


if __name__ == '__main__': 
    app = QApplication(sys.argv) 
    window = MainWindow() 
    window.show() 
    sys.exit(app.exec())
