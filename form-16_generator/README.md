# Sola

## Overview

Sola is a software tool designed to automate the generation of Form-16 documents using ITR format Excel files. It simplifies the process of extracting relevant details and creating professional Form-16 documents.

## Features

- Extracts details from ITR format Excel files.
- Generates Form-16 documents in Excel format.
- Supports additional details like donations, health insurance, and home loan information.
- User-friendly interface for file selection and output generation.

## Installation

1. Clone the repository:

   ```bash
   git clone <repository-url>
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Run the Application

    ```bash
    py ./main.py
    ```

## Usage

1. Launch the application.
2. Select the ITR format file and specify the output path for the Form-16 document.
3. Click "Submit" to generate the Form-16.

## File Structure

- **`scripts/excel_processor.py`**: Contains logic for processing ITR format Excel files.
- **`ui/components.py`**: Implements the user interface.
- **`config/colors.py`**: Defines color schemes for the UI.

## Dependencies

- Python 3.9+
- Flet
- OpenPyXL
- Pandas
