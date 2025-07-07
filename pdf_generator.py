import os
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter


# --- Configuration ---
# IMPORTANT: Place your logo image (e.g., 'logo.png') in the same directory as this script,
# or provide the full, absolute path to the image file here.
LOGO_LEFT_PATH = 'Jims_logo.jpg'
# Add the path for the second logo, which will appear in the center.
LOGO_CENTER_PATH = 'Jims_name.jpg' # Example: place 'second_logo.png' in the same folder


def create_receipt_pdf(data_row: pd.Series, file_path: str):
    """
    Creates a single PDF receipt from a row of data.

    Args:
        data_row (pd.Series): The data for one receipt.
        file_path (str): The full path where the PDF will be saved.

    Returns:
        bool: True if successful, False otherwise.
    """
    try:
        doc = SimpleDocTemplate(file_path, pagesize=letter, rightMargin=inch, leftMargin=inch, topMargin=inch, bottomMargin=inch)
        styles = getSampleStyleSheet()

        flowables = []

        # --- Add Header with Logos ---
        logo_left = Image(LOGO_LEFT_PATH, width=1.4*inch, height=1*inch) if os.path.exists(LOGO_LEFT_PATH) else ""
        logo_center = Image(LOGO_CENTER_PATH, width=3.6*inch, height=1.5*inch) if os.path.exists(LOGO_CENTER_PATH) else ""

        # Use a three-column table for left and center alignment.
        # The third column is an empty placeholder to balance the layout.
        header_data = [[logo_left, logo_center, ""]]

        # Define column widths. The center column takes up the remaining space.
        # The left and right columns act as margins.
        side_width = 2 * inch
        center_width = doc.width - (2 * side_width)

        header_table = Table(header_data, colWidths=[side_width, center_width, side_width])
        header_table.setStyle(TableStyle([
            # Align the left image to the left of its cell
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            # Align the center image to the center of its cell
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            # Vertically align all images to the middle
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        # Only add the header table if at least one logo exists
        if logo_left or logo_center:
            flowables.append(header_table)
            flowables.append(Spacer(1, 0.25 * inch))

        # Title
        title_style = styles['h1']
        title_style.alignment = TA_CENTER
        flowables.append(Paragraph("Fee Receipt", title_style))
        flowables.append(Spacer(1, 0.25 * inch))

        # --- New Table Format ---
        # IMPORTANT: These column names must exactly match the headers in your Excel file.
        receipt_fields = [
            'Name', 'Admission Number', 'Class', 'Bank Reference ID',
            'Order ID', 'Transaction ID', 'Status', 'Amount', 'Date'
        ]

        # Prepare data for the table: a list of [label, value] pairs
        table_data = []
        for field in receipt_fields:
            # Safely get the value, using 'N/A' if the column doesn't exist in the data
            value = data_row.get(field, 'N/A')
            table_data.append([field, str(value)])

        # Create the table with specified column widths
        receipt_table = Table(table_data, colWidths=[2 * inch, 4 * inch])

        # Define and apply a professional table style
        style = TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),          # Left-align the labels (first column)
            ('ALIGN', (1, 0), (1, -1), 'LEFT'),          # Left-align the values (second column)
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'), # Bold the labels
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),      # Add padding to all cells
            ('TOPPADDING', (0, 0), (-1, -1), 8),         # Add padding to all cells
            ('GRID', (0, 0), (-1, -1), 1, colors.black)  # Add a grid to all cells
        ])
        receipt_table.setStyle(style)

        flowables.append(receipt_table)
        # --- End of Table Format ---

        doc.build(flowables)
        return True
    except Exception as e:
        print(f"Error creating PDF {file_path}: {e}")
        return False