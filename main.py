import sys
import os
import subprocess
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QTableWidget, QMessageBox, QFrame,
                             QLineEdit, QGroupBox, QLabel, QHBoxLayout)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QFontMetrics, QFont
from upload_excel import upload_file  # Import the function
from excel_viewer import display_excel_data  # Import the new function
from pdf_generator import create_receipt_pdf # Import the new PDF generator
from table_filter import filter_table_by_name # Import the new filter function
from individual_printer import print_single_receipt_from_df # Import the new individual print logic

def resource_path(relative_path):
    """ Get absolute path to resource for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class MainWindow(QWidget):
    # --- Class Level Configuration ---
    # IMPORTANT: Change this value to match the exact column header for student names in your Excel file.
    STUDENT_NAME_COLUMN = 'Name'

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Admission Reciept 2025")
        self.setMinimumSize(1120, 800)

        # --- State Management ---
        self.name_column_table_index = -1  # The index of the name column in the QTableWidget
        self.df = None
        self.selected_rows = set()

        # --- Widgets ---
        self.upload_button = QPushButton("Upload Excel File")
        self.upload_button.setMinimumHeight(40)  # Increase button height
        self.search_bar = QLineEdit()
        self.search_bar.setMinimumHeight(33)  # Increase search bar height
        self.search_bar.setPlaceholderText("Search by Name...")
        self.print_receipts_button = QPushButton("Print Receipt(s)")
        self.print_receipts_button.setMinimumHeight(40)  # Increase button height

        # --- Set Button Font ---
        button_font = QFont()
        button_font.setPointSize(11) # Set a larger font size for the button text
        self.upload_button.setFont(button_font)
        self.print_receipts_button.setFont(button_font)

        # --- Set Button Style ---
        button_style = "background-color: #2b2e9b; color: white; border-radius: 5px;"
        self.upload_button.setStyleSheet(button_style)
        self.print_receipts_button.setStyleSheet(button_style)

        self.table_widget = QTableWidget()
        # Make the table read-only to prevent accidental edits
        self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)
        # Add padding to all data cells for better readability
        self.table_widget.setStyleSheet("QTableWidget::item { padding: 8px; }")

        # --- Layout ---
        main_layout = QVBoxLayout(self)

        # --- Logo Container ---
        logo_container = QWidget()
        logo_layout = QHBoxLayout(logo_container)
        logo_layout.setContentsMargins(10, 10, 10, 10) # Add padding on all sides

        # --- Left Logo (Jims_logo.jpg) ---
        logo_left_label = QLabel()
        logo_left_path = resource_path('Jims_logo-removebg-preview.png')
        pixmap_left = None
        if os.path.exists(logo_left_path):
            pixmap_left = QPixmap(logo_left_path)
            # Scale to a height to keep proportions, as it's more square
            logo_left_label.setPixmap(pixmap_left.scaledToHeight(125, Qt.SmoothTransformation))
            logo_left_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        # --- Center Logo (Jims_name.jpg) ---
        logo_center_label = QLabel()
        logo_center_path = resource_path('Jims_name-removebg-preview.png')
        if os.path.exists(logo_center_path):
            pixmap_center = QPixmap(logo_center_path)
            # Define a fixed height for the logo, matching the left logo for alignment.
            logo_height = 110
            # Scale to a fixed width and height, ignoring the aspect ratio.
            # This will distort the image if the new dimensions don't match the original ratio.
            logo_center_label.setPixmap(pixmap_center.scaled(500, logo_height, Qt.IgnoreAspectRatio, Qt.SmoothTransformation))
            logo_center_label.setAlignment(Qt.AlignCenter | Qt.AlignTop)

        # Use a dummy widget on the right to balance the left logo, ensuring the center logo is truly centered.
        dummy_widget = QWidget()
        if pixmap_left:
            # The dummy widget must have the same width as the visible pixmap on the left label.
            dummy_widget.setFixedWidth(80)

        # Add widgets to the layout to achieve the desired alignment.
        logo_layout.addWidget(logo_left_label)
        logo_layout.addStretch()
        logo_layout.addWidget(logo_center_label)
        logo_layout.addStretch()
        logo_layout.addWidget(dummy_widget)

        # Add the container to the main layout
        main_layout.addWidget(logo_container)

        # --- Horizontal Line Separator ---
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken) # This provides a 3D sunken line, which is typically grey
        main_layout.addWidget(line)

        # Add some vertical space between the line and the title
        main_layout.addSpacing(10)

        # --- Title Label ---
        title_label = QLabel("Admission Reciept 2025")
        title_label.setAlignment(Qt.AlignCenter)
        # Make the font bigger and bold for emphasis
        font = title_label.font()
        font.setPointSize(16)
        font.setBold(True)
        title_label.setFont(font)
        main_layout.addWidget(title_label)

        # --- Set Button Width to Match Title ---
        # Calculate the pixel width of the title label to make the button match
        fm = QFontMetrics(font)
        text_width = fm.width("Admission Reciept 2025")
        # Add some horizontal padding for a better look
        button_width = text_width + 40
        self.upload_button.setFixedWidth(button_width)
        self.print_receipts_button.setFixedWidth(button_width)

        # Add some vertical space between the title and the upload button
        main_layout.addSpacing(15)

        # --- Upload Section ---
        # Use a container to place a label next to the button
        upload_container = QWidget()
        upload_layout = QHBoxLayout(upload_container)
        upload_layout.setContentsMargins(0, 0, 0, 0) # Remove layout's own margins

        upload_layout.addStretch() # Add stretch before to start centering
        upload_layout.addWidget(self.upload_button)
        upload_layout.addStretch() # Add stretch after to finish centering

        main_layout.addWidget(upload_container)

        # Use a QGroupBox for a visually and structurally robust container
        table_group_box = QGroupBox("Data Table")
        table_layout = QVBoxLayout(table_group_box)
        table_layout.addWidget(self.search_bar)
        table_layout.addWidget(self.table_widget, 1)

        # Add the group box to the main layout with a stretch factor.
        # This makes the entire table area expand and shrink with the window.
        main_layout.addWidget(table_group_box, 1)

        # --- Print Section ---
        print_container = QWidget()
        print_layout = QHBoxLayout(print_container)
        print_layout.setContentsMargins(0, 0, 0, 0)
        print_layout.addStretch()
        print_layout.addWidget(self.print_receipts_button)
        print_layout.addStretch()
        main_layout.addWidget(print_container)

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
        Sends a file to the default printer. If direct printing fails,
        it logs the error to the terminal and opens the file for manual printing.
        """
        try:
            if sys.platform == "win32":
                # Use the default Windows shell command for printing.
                os.startfile(filepath, "print")
            elif sys.platform == "darwin":  # macOS
                # lp is the standard printing command on macOS.
                subprocess.run(["lp", filepath], check=True)
            elif sys.platform.startswith("linux"):  # Linux
                subprocess.run(["lp", filepath], check=True)
            else:
                raise NotImplementedError(f"Automatic printing not supported on {sys.platform}")

        except Exception as e:
            # Automatic printing failed. Log the error to the terminal.
            print(f"PRINTING_ERROR: Could not print '{os.path.basename(filepath)}' automatically. Reason: {e}", file=sys.stderr)
            print("INFO: Fallback - attempting to open file for manual printing.", file=sys.stderr)
            try:
                # Cross-platform fallback to open the file in the default viewer.
                if sys.platform == "win32":
                    os.startfile(filepath)
                elif sys.platform == "darwin":
                    subprocess.run(["open", filepath], check=True)
                else: # linux and other unix
                    subprocess.run(["xdg-open", filepath], check=True)
            except Exception as open_e:
                # If opening the file also fails, log that critical error to the terminal.
                print(f"CRITICAL_ERROR: Could not open '{filepath}' for manual printing. Reason: {open_e}", file=sys.stderr)

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
        print(f"PDF Generation Complete. Success: {success_count}, Failed: {error_count}")

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
