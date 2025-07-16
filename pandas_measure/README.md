# XLSX Generation Performance Test

This script provides a simple way to test the performance of generating large `.xlsx` (Excel) files using the openpyxl library in Python. It measures the total time taken, the peak memory consumed by the process, and the final size of the generated file.

## Documentation Optimised Modes
[OpenPyXL - Optimised Modes](https://openpyxl.readthedocs.io/en/stable/optimized.html)

## Description
The script contains a primary function, create_large_xlsx, that generates an Excel workbook with a configurable number of rows and columns. To optimize for large datasets, it leverages openpyxl's write_only mode, which is significantly more memory-efficient than the standard mode.

The main execution block of the script wraps the file creation process with timers and memory usage monitors from the `time` and `psutil` libraries, respectively, to provide clear performance metrics upon completion.

## Requirements
To run this script, you need to have Python installed, along with the following libraries:

 - `openpyxl`: For creating and manipulating .xlsx files.

- `psutil`: For accessing system details and process utilities, used here to monitor memory usage.

You can install these dependencies using pip:

```
pip install openpyxl lxml psutil
```

## Usage

1. Save the code as a Python file (e.g., `performance_test.py`).

2. Open your terminal or command prompt.

3. Navigate to the directory where you saved the file.

4. Run the script using the following command:

## How It Works
Initial Snapshot: Before starting the file creation, the script records the current time and the process's initial memory usage.

File Generation: It proceeds to create the .xlsx file, writing a header and then iterating through the specified number of rows to append data. Progress is printed every 10,000 rows.

Final Snapshot: After the file is saved, the script captures the end time and the final memory usage.

Reporting: It calculates the total elapsed time and the increase in memory usage (initial vs. final) to report the performance metrics. It also reports the size of the output file.

The script will print its progress to the console and display a summary of the performance results once the file has been created.

## How It Works

1. **Initial Snapshot**: Before starting the file creation, the script records the current time and the process's initial memory usage.

2. **File Generation**: It proceeds to create the `.xlsx` file, writing a header and then iterating through the specified number of rows to append data. Progress is printed every 10,000 rows.

3. **Final Snapshot**: After the file is saved, the script captures the end time and the final memory usage.

4. **Reporting**: It calculates the total elapsed time and the increase in memory usage (initial vs. final) to report the performance metrics. It also reports the size of the output file.

## Customization

You can easily change the size of the spreadsheet to test different scenarios. In the `if __name__ == "__main__":` block, modify the following variables:

* `NUM_ROWS`: The number of data rows to generate.

* `NUM_COLS`: The number of columns for each row.

* `OUTPUT_FILENAME`: The name of the output `.xlsx` file.

For example, to create a smaller file:

## Define the size of the spreadsheet
```
NUM_ROWS = 5000
NUM_COLS = 10
OUTPUT_FILENAME = "small_test_file.xlsx"
```
