# Point List Generator

Python tool for CAISO points list generation. 

This application reads analog and digital point lists (AI, AO, DI, DO) and automatically produces a formatted Excel workbook used for RTAC configuration and SCADA integration.

The tool includes a PyQt GUI for selecting input files and generating the output workbook.

## Features

* Supports Analog Inputs, Analog Outputs, Digital Inputs, and Digital Outputs
* Automatic CAISO and Substation tag generation
* Generates structured Excel workbook with:

  * Point Selection sheet
  * CAISO configuration sheet
  * Substation configuration sheet
  * Meter-specific sheets
* Meter filtering based on DNP indices
* Clean Excel formatting with tables and section headers
* GUI interface built with PyQt

## Project Structure

```
app.py                 # Application entry point
main_window.py         # PyQt GUI
worker.py              # Background processing thread

io_reader.py           # Input file reading and header detection

transform_common.py    # Shared transformation utilities
transform_analog.py    # AI / AO transformation logic
transform_digital.py   # DI / DO transformation logic

excel_writer.py        # Workbook creation
excel_renderers.py     # Sheet rendering
excel_utils.py         # Excel helpers
excel_styles.py        # Excel styling
```

## Installation

Clone the repository:

```
git clone https://github.com/YOUR_USERNAME/rtac-pointlist-generator.git
cd rtac-pointlist-generator
```

Install dependencies:

```
pip install -r requirements.txt
```

## Running the Application

Start the GUI:

```
python app.py
```

Then:

1. Select AI / AO / DI / DO input files
2. Configure meter definitions (optional)
3. Choose output Excel file
4. Click **Generate Excel**

The program will produce a formatted workbook containing all RTAC configuration sheets.

## Requirements

* Python 3.9+
* pandas
* openpyxl
* PyQt6 (or PyQt5 fallback)

## Notes

* Input spreadsheets must contain the required columns such as `Point Description`, `Point Name`, and `DNP Index`.
* The application automatically detects header rows in exported spreadsheets.
* DNP indices are preserved (no renumbering or compression).

## License

Internal engineering tool. Use as needed for RTAC workflow automation.
