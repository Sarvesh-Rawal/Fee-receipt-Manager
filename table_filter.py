from PyQt5.QtWidgets import QTableWidget


def filter_table_by_name(table_widget: QTableWidget, search_text: str, name_column_index: int):
    """
    Filters the rows of a QTableWidget based on a search text in a specific column.
    The search is case-insensitive.

    Args:
        table_widget (QTableWidget): The table to filter.
        search_text (str): The text to search for.
        name_column_index (int): The index of the column to search in.
    """
    search_text = search_text.lower()
    for row_idx in range(table_widget.rowCount()):
        item = table_widget.item(row_idx, name_column_index)
        # Make sure the item exists before trying to access its text
        if item and search_text in item.text().lower():
            table_widget.setRowHidden(row_idx, False)
        else:
            # Hide the row if the item doesn't exist or doesn't match
            table_widget.setRowHidden(row_idx, True)