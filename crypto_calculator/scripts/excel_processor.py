import openpyxl.cell
import pandas as pd
from dataclasses import dataclass
from math import isnan
import random as r
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


@dataclass
class ExcelProcessor:
    def _select_workbook(self, file_name: str, sheet_name: str = "FORM-16") -> None:
        """
        Opens an existing Excel workbook for Form-16 using openpyxl and selects the specified worksheet.

        This method loads the Excel file provided by `file_name` and sets the `self.workbook` attribute
        to the loaded workbook object. It then selects the worksheet specified by `sheet_name` (defaulting
        to "FORM-16") and assigns it to the `self.ws` attribute for further processing.

        Args:
            file_name (str): The path to the Excel file to open. This should be a valid .xlsx file.
            sheet_name (str, optional): The name of the worksheet to select from the workbook.
            Defaults to "FORM-16".

        Raises:
            FileNotFoundError: If the specified Excel file does not exist.
            KeyError: If the specified worksheet name does not exist in the workbook.

        Side Effects:
            Sets self.workbook to the loaded openpyxl workbook.
            Sets self.ws to the selected worksheet within the workbook.
        """
        self.workbook = openpyxl.load_workbook(file_name)
        self._select_worksheet(self.workbook, sheet_name)

    def _select_worksheet(self, workbook: openpyxl.Workbook, sheet_name: str) -> None:
        """
        Selects a worksheet from an openpyxl workbook by name and assigns it to self.ws.

        This method takes an openpyxl Workbook object and a worksheet name, and sets the
        self.ws attribute to the corresponding worksheet. It is used internally after loading
        a workbook to prepare for further processing or data manipulation.

        Args:
            workbook (openpyxl.Workbook): The loaded Excel workbook object.
            sheet_name (str): The name of the worksheet to select.

        Raises:
            KeyError: If the specified worksheet name does not exist in the workbook.

        Side Effects:
            Sets self.ws to the selected worksheet within the workbook.
        """
        try:
            self.ws = self.workbook[sheet_name]
            # Sheet exists, use it
        except KeyError:
            # Sheet doesn't exist
            self.ws = self.workbook.create_sheet(sheet_name, 0)

    def _apply_style(self, cell, style_key: str) -> None:
        """
        Applies a predefined style from self.s to a given cell.

        Args:
            cell (openpyxl.cell.cell.Cell): The cell to apply the style to.
            style_key (str): The key of the style in self.s to apply.

        Raises:
            KeyError: If the style_key does not exist in self.s.
        """
        if style_key not in self.s:
            raise KeyError(f"Style '{style_key}' not found in self.s")

        style = self.s[style_key]
        for attribute, value in style.items():
            setattr(cell, attribute, value)

    def _add_formats(self) -> None:
        """
        Defines commonly used cell styles for formatting Excel sheets and stores them in self.s.

        Side Effects:
            Sets self.s to a dictionary containing predefined cell styles.
        """
        # Define font styles

        bold_font = Font(name="calibri", bold=True, size=16)
        blue_font = Font(name="calibri", size=18, bold=True, color="FFFFFF")
        normal_font = Font(name="calibri", size=12)
        big_font = Font(name="calibri", bold=True, size=26)
        black_font = Font(name="calibri", color="FFFFFF", bold=True, size=16)
        black_font_h = Font(name="calibri", color="FFFFFF", bold=True, size=26)

        # Define alignment styles
        center_alignment = Alignment(
            horizontal="center", vertical="center", wrapText=True
        )

        # Define border styles
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        # Define fill styles for background colors
        dark_red_fill = PatternFill(
            start_color="FF0066", end_color="FF0066", fill_type="solid"
        )
        medium_red_fill = PatternFill(
            start_color="FF3399", end_color="FF3399", fill_type="solid"
        )
        light_red_fill = PatternFill(
            start_color="FF6699", end_color="FF6699", fill_type="solid"
        )
        black_fill = PatternFill(
            start_color="262626", end_color="262626", fill_type="solid"
        )
        blue_fill = PatternFill(
            start_color="002060", end_color="002060", fill_type="solid"
        )
        green_fill = PatternFill(
            start_color="A9D08E", end_color="A9D08E", fill_type="solid"
        )

        # Store styles in self.s
        self.s = {
            "blank": {"font": normal_font, "alignment": center_alignment},
            "dark_red": {
                "font": bold_font,
                "alignment": center_alignment,
                "fill": dark_red_fill,
                "border": thin_border,
            },
            "medium_red": {
                "font": bold_font,
                "alignment": center_alignment,
                "fill": medium_red_fill,
                "border": thin_border,
            },
            "light_red": {
                "font": bold_font,
                "alignment": center_alignment,
                "fill": light_red_fill,
                "border": thin_border,
            },
            "black": {
                "font": black_font,
                "alignment": center_alignment,
                "fill": black_fill,
                "border": thin_border,
            },
            "blue": {
                "font": blue_font,
                "alignment": center_alignment,
                "fill": blue_fill,
                "border": thin_border,
            },
            "green": {
                "font": normal_font,
                "alignment": center_alignment,
                "fill": green_fill,
                "border": thin_border,
            },
            "green_h": {
                "font": big_font,
                "alignment": center_alignment,
                "fill": green_fill,
                "border": thin_border,
            },
            "dark_red_h": {
                "font": big_font,
                "alignment": center_alignment,
                "fill": dark_red_fill,
                "border": thin_border,
            },
            "medium_red_h": {
                "font": bold_font,
                "alignment": center_alignment,
                "fill": medium_red_fill,
                "border": thin_border,
            },
            "light_red_h": {
                "font": big_font,
                "alignment": center_alignment,
                "fill": light_red_fill,
                "border": thin_border,
            },
            "black_h": {
                "font": black_font_h,
                "alignment": center_alignment,
                "fill": black_fill,
                "border": thin_border,
            },
        }

    def _set_worksheet_dimensions(self):
        """
        Set all rows to height 43 and all columns to width 26.1
        """
        from openpyxl.utils import get_column_letter

        # Set row heights for all existing rows plus extra buffer
        max_row = max(self.ws.max_row, 100)  # At least 100 rows
        for row in range(1, max_row + 1):
            self.ws.row_dimensions[row].height = 43

        # Set column widths for all existing columns plus extra buffer
        max_col = max(self.ws.max_column, 26)  # At least 26 columns (A-Z)
        for col in range(1, max_col + 1):
            col_letter = get_column_letter(col)
            self.ws.column_dimensions[col_letter].width = (
                26.1 + 0.71
            )  # type: ignore # Some Randome shit decreases the width by 0.71

        # Set defaults for any new rows/columns that might be added later
        self.ws.sheet_format.defaultRowHeight = 43
        self.ws.sheet_format.defaultColWidth = (
            26.1 + 0.71
        )  # Some Randome shit decreases the width by 0.71
        self.ws.sheet_format.customHeight = False  # Don't auto-adjust heights

    def _set_default_style(self, max_rows: int = 200, max_cols: int = 50) -> None:
        """
        Apply default 'blank' style to all cells in the worksheet
        Args:
            max_rows: Number of rows to style (default 200)
            max_cols: Number of columns to style (default 50)
        """
        for row in range(1, max_rows + 1):
            for col in range(1, max_cols + 1):
                cell = self.ws.cell(row=row, column=col)
                self._apply_style(cell, "blank")

    def set(
        self, cell: openpyxl.cell.Cell, value: str | int | float, style: str
    ) -> None:
        """
        Sets the cell with value and style
        """
        cell.value = value
        self._apply_style(cell, style)

    def make_dashboard(self, file_name: str, sheet_name: str = "crypto") -> None:
        """
        Creates the Dashboard for crypto Calculations
        """

        self._select_workbook(file_name, sheet_name)

        self._add_formats()

        self._set_worksheet_dimensions()
        self._set_default_style()

        self.ws.merge_cells("A1:J1")
        crypto_details = self.ws["A1"]
        self.set(crypto_details, "Crypto Details", "black_h")

        self.ws.merge_cells("A2:B2")
        source = self.ws["A2"]
        self.set(source, "Source", "dark_red")

        doa = self.ws["C2"]
        self.set(doa, "Date of Acquisition", "medium_red")

        dot = self.ws["D2"]
        self.set(dot, "Date of Transfer", "light_red")

        coa = self.ws["E2"]
        self.set(coa, "Cost of Acquisition", "dark_red")

        cr = self.ws["F2"]
        self.set(cr, "Consideration Recived", "medium_red")

        cg = self.ws["g2"]
        self.set(cg, "Capital Gain", "blue")

        self.ws.merge_cells("H2:J2")
        msg = self.ws["H2"]
        self.set(msg, "Accha Khasa LOSS Hua Hai Bhai Saab", "black")

        """################# Total Capital Gain #################"""

        self.ws.merge_cells("H3:J3")
        tcg_title = self.ws["H3"]
        self.set(tcg_title, "Total Capital Gain", "blue")

        self.ws.merge_cells("H4:J7")
        tcg_value = self.ws["H4"]
        self.set(tcg_value, r.randint(10000, 100000), "dark_red_h")

        """################# Total Cost of Acquisition #################"""

        self.ws.merge_cells("H8:J8")
        tcoa_title = self.ws["H8"]
        self.set(tcoa_title, "Total Cost of Acquisition", "blue")

        self.ws.merge_cells("H9:J10")
        tcoa_value = self.ws["H9"]
        self.set(tcoa_value, r.randint(10000, 100000), "light_red_h")

        """################# Total Consideration Recived #################"""

        self.ws.merge_cells("H11:J11")
        tcr_title = self.ws["H11"]
        self.set(tcr_title, "Total Consideration Recived", "blue")

        self.ws.merge_cells("H12:J14")
        tcr_value = self.ws["H12"]
        self.set(tcr_value, r.randint(10000, 100000), "green_h")

        self.workbook.close()
        self.workbook.save(file_name)


if __name__ == "__main__":
    # Create an instance of CSVProcessor
    test = ExcelProcessor()

    test.make_dashboard(
        file_name="test/Form-16 .xlsx",
    )
