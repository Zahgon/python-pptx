"""Chart builder and related objects."""

from __future__ import annotations

import io
from contextlib import contextmanager

from xlsxwriter import Workbook


class _BaseWorkbookWriter(object):
    """Base class for workbook writers, providing shared members."""

    def __init__(self, chart_data):
        super(_BaseWorkbookWriter, self).__init__()
        self._chart_data = chart_data

    @property
    def xlsx_blob(self):
        """bytes for Excel file containing chart_data."""
        pass

    @contextmanager
    def _open_worksheet(self, xlsx_file):
        """
        Enable XlsxWriter Worksheet object to be opened, operated on, and
        then automatically closed within a `with` statement. A filename or
        stream object (such as an `io.BytesIO` instance) is expected as
        *xlsx_file*.
        """
        pass

    def _populate_worksheet(self, workbook, worksheet):
        """
        Must be overridden by each subclass to provide the particulars of
        writing the spreadsheet data.
        """
        raise NotImplementedError("must be provided by each subclass")


class CategoryWorkbookWriter(_BaseWorkbookWriter):
    """
    Determines Excel worksheet layout and can write an Excel workbook from
    a CategoryChartData object. Serves as the authority for Excel worksheet
    ranges.
    """

    @property
    def categories_ref(self):
        """
        The Excel worksheet reference to the categories for this chart (not
        including the column heading).
        """
        pass

    def series_name_ref(self, series):
        """
        Return the Excel worksheet reference to the cell containing the name
        for *series*. This also serves as the column heading for the series
        values.
        """
        pass

    def values_ref(self, series):
        """
        The Excel worksheet reference to the values for this series (not
        including the column heading).
        """
        pass

    @staticmethod
    def _column_reference(column_number):
        """Return str Excel column reference like 'BQ' for *column_number*.

        *column_number* is an int in the range 1-16384 inclusive, where
        1 maps to column 'A'.
        """
        pass

    def _populate_worksheet(self, workbook, worksheet):
        """
        Write the chart data contents to *worksheet* in category chart
        layout. Write categories starting in the first column starting in
        the second row, and proceeding one column per category level (for
        charts having multi-level categories). Write series as columns
        starting in the next following column, placing the series title in
        the first cell.
        """
        pass

    def _series_col_letter(self, series):
        """
        The letter of the Excel worksheet column in which the data for a
        series appears.
        """
        pass

    def _write_categories(self, workbook, worksheet):
        """
        Write the categories column(s) to *worksheet*. Categories start in
        the first column starting in the second row, and proceeding one
        column per category level (for charts having multi-level categories).
        A date category is formatted as a date. All others are formatted
        `General`.
        """
        pass

    def _write_cat_column(self, worksheet, col, level, num_format):
        """
        Write a category column defined by *level* to *worksheet* at offset
        *col* and formatted with *num_format*.
        """
        pass

    def _write_series(self, workbook, worksheet):
        """
        Write the series column(s) to *worksheet*. Series start in the column
        following the last categories column, placing the series title in the
        first cell.
        """
        pass


class XyWorkbookWriter(_BaseWorkbookWriter):
    """
    Determines Excel worksheet layout and can write an Excel workbook from XY
    chart data. Serves as the authority for Excel worksheet ranges.
    """

    def series_name_ref(self, series):
        """
        Return the Excel worksheet reference to the cell containing the name
        for *series*. This also serves as the column heading for the series
        Y values.
        """
        pass

    def series_table_row_offset(self, series):
        """
        Return the number of rows preceding the data table for *series* in
        the Excel worksheet.
        """
        pass

    def x_values_ref(self, series):
        """
        The Excel worksheet reference to the X values for this chart (not
        including the column label).
        """
        pass

    def y_values_ref(self, series):
        """
        The Excel worksheet reference to the Y values for this chart (not
        including the column label).
        """
        pass

    def _populate_worksheet(self, workbook, worksheet):
        """
        Write chart data contents to *worksheet* in the standard XY chart
        layout. Write the data for each series to a separate two-column
        table, X values in column A and Y values in column B. Place the
        series label in the first (heading) cell of the column.
        """
        pass


class BubbleWorkbookWriter(XyWorkbookWriter):
    """
    Service object that knows how to write an Excel workbook from bubble
    chart data.
    """

    def bubble_sizes_ref(self, series):
        """
        The Excel worksheet reference to the range containing the bubble
        sizes for *series* (not including the column heading cell).
        """
        pass

    def _populate_worksheet(self, workbook, worksheet):
        """
        Write chart data contents to *worksheet* in the bubble chart layout.
        Write the data for each series to a separate three-column table with
        X values in column A, Y values in column B, and bubble sizes in
        column C. Place the series label in the first (heading) cell of the
        values column.
        """
        pass
