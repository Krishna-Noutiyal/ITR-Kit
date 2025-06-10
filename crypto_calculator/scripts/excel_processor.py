import pandas as pd
from dataclasses import dataclass
from math import isnan
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
        self.ws = workbook[sheet_name]


def _add_formats(self) -> None:
    """
    Defines commonly used cell styles for formatting Excel sheets and stores them in self.s.

    Side Effects:
        Sets self.s to a dictionary containing predefined cell styles.
    """
    # Define font styles
    bold_font = Font(name="Calibri", bold=True, size=16)

    # Define alignment styles
    center_alignment = Alignment(horizontal="center", vertical="center")

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
    blue_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")

    # Store styles in self.s
    self.s = {
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
            "font": bold_font,
            "alignment": center_alignment,
            "fill": black_fill,
            "border": thin_border,
        },
        "blue": {
            "font": bold_font,
            "alignment": center_alignment,
            "fill": blue_fill,
            "border": thin_border,
        },
    }
    def make_dashboard(self, file_name:str, sheet_name: str = "crypto") -> None:
        """
        Creates the Dashboard for crypto Calculations
        """
        pass


if __name__ == "__main__":
    # Create an instance of CSVProcessor
    test = ExcelProcessor()

    test.create_form_16(
        itr_format="form-16_generator/test/ITR Format (PIC).xlsx",
        form_16="form-16_generator/test/Form-16.xlsx",
    )
