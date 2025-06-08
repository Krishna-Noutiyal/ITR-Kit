import glob, os
import pandas as pd
from dataclasses import dataclass
import xlsxwriter
import openpyxl
import xlsxwriter.worksheet


@dataclass
class ExcelProcessor:
    def _select_form16(self, file_name: str, sheet_name: str = "FORM-16") -> None:
        """
        Open an existing Excel workbook for Form-16 and select the worksheet using openpyxl.

        Args:
            file_name (str): The name of the Excel file to open.
            sheet_name (str, optional): The name of the worksheet to select. Defaults to "FORM-16".
        """
        self.form16 = openpyxl.load_workbook(file_name)
        self.ws = self.form16[sheet_name]

    def _extract_details(self, file_path: str, sheet_name: str = "ITR Format") -> dict:
        """
        Extract details from the ITR format worksheet using pandas.

        Args:
            file_path (str): The path to the Excel file.
            sheet_name (str, optional): The name of the worksheet to extract from.
        Returns:
            dict: A dictionary containing extracted details.
        """
        self.data = {}
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        # Extract key-value pairs from B5:C20 (Excel is 1-indexed, pandas is 0-indexed)
        for row in range(4, 20):  # B5 is row 4 (0-indexed), C20 is row 19
            key = df.iat[row, 1]  # Column B (0:Serial_No, so index 1)
            value = df.iat[row, 2]  # Column C (0:Serial_No, so index 2)

            # Only add to dictionary if both key and value are not NaN
            if pd.notna(key) and pd.notna(value):
                self.data[key] = value

        # Extract key-list pairs from B22:D51 (Excel is 1-indexed, pandas is 0-indexed)
        for row in range(21, 51):  # B22 is row 21, D51 is row 50
            key = df.iat[row, 1]  # Column B (index 1)
            val1 = df.iat[row, 2]  # Column C (index 2)
            val2 = df.iat[row, 3]  # Column D (index 3)
            if pd.notna(key) and (pd.notna(val1) or pd.notna(val2)):
                self.data[key] = [val1, val2]

        self.data["passwd"] = df.iat[
            12, 3
        ]  # D13 is row 12 (0-indexed), column 3 (0-indexed)
        for key, value in self.data.items():
            print(f"{key} : {value}")

        df = pd.read_excel(file_path, sheet_name="Home Loan", header=None)
        
        self.data["bank_name"] = df.iat[3, 1]
        self.data["loan_ac_number"] = df.iat[3, 2]  
        self.data["date_of_sanction"] = df.iat[3, 3]  
        self.data["total_loan_amount"] = df.iat[3, 4]
        self.data["bank_name2"] = df.iat[4, 1]
        self.data["loan_ac_number2"] = df.iat[4, 2]  
        self.data["date_of_sanction2"] = df.iat[4, 3]  
        self.data["total_loan_amount2"] = df.iat[4, 4]

        return self.data

    def create_form_16(self, itr_format: str, form_16: str) -> bool:
        """
        Create Form-16 from the given ITR format file.
        """
        try:
            self._select_form16(form_16)
            details = self._extract_details(itr_format)

            self.form16.save(form_16)
            self.form16.close()

            return True
        except Exception as e:
            print(f"Error creating Form-16: {e}")
            return False


if __name__ == "__main__":
    # Create an instance of CSVProcessor
    test = ExcelProcessor()

    test.create_form_16(
        itr_format="form-16_generator/test/ITR Format (PIC).xlsx",
        form_16="form-16_generator/test/Form-16.xlsx",
    )
