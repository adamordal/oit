# OIT Billing Script

## Overview

This project contains a set of Python scripts designed to process quota usage data from JSON configuration files and generate a report in an Excel template. The main functionalities include reading data from Excel and JSON files, calculating totals for different departments, and writing the results to an Excel template.

## Files

### `main.py`

This is the main script that orchestrates the entire process. It performs the following tasks:
- Prompts the user to select the Ingram Micro cost report file.
- Loads the Ingram Micro cost report data.
- Prompts the user to select the LWX and LWN JSON configuration files.
- Loads the JSON configuration files and extracts quota usage data.
- Calculates totals for LWX and LWN departments.
- Prompts the user to specify the output file name.
- Writes the data to the Excel template.

### `extract_quota_usage.py`

This script contains functions to extract quota usage data from a JSON configuration file. It includes:
- `add_parser_options(parser)`: Adds command line options for JSON file and encoding.
- `extract_quota_usage(json_cfg)`: Extracts relevant fields for "Quota Usage" from the JSON config.
- `main()`: Sets up logging, parses command line options, loads the JSON configuration file, and extracts quota usage data.

### `logging_setup.py`

This script sets up the logging configuration. It includes:
- `setup_logging()`: Configures the logging settings and returns the logger object.

### `file_operations.py`

This script contains functions for file operations such as reading from Excel files, browsing for files, saving files, and writing data to an Excel template. It includes:
- `read_xlsx_to_dict(file_path, sheet_name, key_column, value_start_column)`: Reads data from an Excel file and returns it as a dictionary.
- `browse_file(prompt, filetypes)`: Opens a file dialog to browse for a file.
- `save_file(prompt)`: Opens a file dialog to save a file.
- `write_to_template(data_dict_lwx, data_dict_lwn, LWX_Cost, LWN_Cost, output_path, departments_lwx, departments_lwn)`: Writes the usage data to the Excel template.

## Input Files

### Ingram Micro Cost Report File

- **Format**: Excel (.xlsx)
- **Sheet Name**: `Rating Report`
- **Columns**: The script reads data starting from the second row and uses the first column as keys and subsequent columns as values.

### LWX and LWN JSON Configuration Files

- **Format**: JSON (.json)
- **Structure**: The JSON files should contain quota usage data under the `stats -> smartquotas -> usage` path.

## Usage

1. Run the `main.py` script.
2. Follow the prompts to select the Ingram Micro cost report file, LWX JSON configuration file, and LWN JSON configuration file.
3. Specify the output file name for the generated report.
4. The script will process the data and save the results to the specified output file.

## Example

```sh
python main.py
```

Follow the prompts to select the required files and specify the output file name.

## Dependencies

- `openpyxl`
- `tkinter`
- `json`
- `logging`
- `optparse`