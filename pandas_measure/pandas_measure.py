from pathlib import Path
import pandas as pd
import numpy as np
import time
import psutil
import os

def get_memory_usage_mb():
    """Returns the memory usage of the current process in megabytes."""
    process = psutil.Process(os.getpid())
    # rss: Resident Set Size is the non-swapped physical memory a process has used.
    mem_info = process.memory_info().rss
    return mem_info / (1024 * 1024)

def create_large_xlsx_with_pandas(filename="large_spreadsheet_pandas.xlsx", rows=100000, cols=50):
    """
    Creates a large XLSX file using pandas and NumPy.
    
    Args:
        filename (str): The name of the file to create.
        rows (int): The number of data rows to generate.
        cols (int): The number of data columns to generate.
    """
    print(f"Starting XLSX creation for {rows} rows and {cols} columns using pandas...")
    
    # --- Generate data efficiently using NumPy ---
    # This is much faster than creating data in Python loops.
    print("  ... Generating data in memory...")
    data = np.random.rand(rows, cols)
    
    # --- Create a list of column names ---
    columns = [f"Column {i+1}" for i in range(cols)]
    
    # --- Create a pandas DataFrame ---
    # The DataFrame holds all the data in memory before writing.
    print("  ... Creating pandas DataFrame...")
    df = pd.DataFrame(data, columns=columns)
    
    # --- Save the DataFrame to an Excel file ---
    # The to_excel method handles the creation and writing of the .xlsx file.
    # engine='openpyxl' is required for .xlsx format.
    # index=False prevents pandas from writing the DataFrame index as a column.
    print(f"  ... Saving DataFrame to '{filename}'...")
    df.to_excel(filename, index=False, engine='openpyxl')
    
    print("Workbook saved successfully.")


if __name__ == "__main__":
    # --- Performance Measurement Setup ---
    start_mem = get_memory_usage_mb()
    start_time = time.perf_counter()

    print("--- Pandas Performance Test Start ---")
    print(f"Initial Memory Usage: {start_mem:.2f} MB")
    print("-" * 35)

    # --- Main Task ---
    # Define the size of the spreadsheet
    NUM_ROWS = 100000
    NUM_COLS = 50
    OUTPUT_FILENAME = "performance_test_pandas.xlsx"
    Path(OUTPUT_FILENAME).unlink(True)

    # Check if the file exists and remove it to ensure a fresh run
    if os.path.exists(OUTPUT_FILENAME):
        os.remove(OUTPUT_FILENAME)

    create_large_xlsx_with_pandas(filename=OUTPUT_FILENAME, rows=NUM_ROWS, cols=NUM_COLS)

    # --- Performance Measurement End ---
    end_time = time.perf_counter()
    end_mem = get_memory_usage_mb()

    elapsed_time = end_time - start_time
    # Note: With pandas, the entire dataset is loaded into memory, 
    # so the peak memory will be significantly higher.
    peak_mem_usage = end_mem - start_mem
    file_size = os.path.getsize(OUTPUT_FILENAME) / (1024 * 1024)

    print("-" * 35)
    print("--- Pandas Performance Test Results ---")
    print(f"Process finished in: {elapsed_time:.4f} seconds")
    print(f"Peak memory usage during process: {peak_mem_usage:.2f} MB")
    print(f"Final memory usage: {end_mem:.2f} MB")
    print(f"Final file size: {file_size:.2f} MB")
    print("-" * 35)
    
    # Instructions to run
    print("\nTo run this script:")
    print("1. Make sure you have the required libraries: pip install pandas numpy openpyxl psutil")
    print(f"2. Run the script from your terminal: python {os.path.basename(__file__)}")
