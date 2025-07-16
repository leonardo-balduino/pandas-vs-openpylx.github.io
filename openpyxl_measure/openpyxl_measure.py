from pathlib import Path
import openpyxl
import time
import psutil
import os
import random

def get_memory_usage_mb():
    """Returns the memory usage of the current process in megabytes."""
    process = psutil.Process(os.getpid())
    # rss: Resident Set Size is the non-swapped physical memory a process has used.
    mem_info = process.memory_info().rss
    return mem_info / (1024 * 1024)

def create_large_xlsx(filename="large_spreadsheet.xlsx", rows=100000, cols=20, write_only=True):
    """
    Creates a large XLSX file with the specified number of rows and columns.
    
    Args:
        filename (str): The name of the file to create.
        rows (int): The number of data rows to generate.
        cols (int): The number of data columns to generate.
    """    
    print(f"Starting XLSX creation for {rows} rows and {cols} columns...")
    
    # --- Create a new workbook and select the active sheet ---
    # write_only=True mode is highly recommended for writing large files.
    # It is much more memory-efficient as it writes data directly to the disk.
    workbook = openpyxl.Workbook(write_only=write_only)
    worksheet = workbook.create_sheet("Performance Test Data")

    # --- Create a header row ---
    header = [f"Column {i+1}" for i in range(cols)]
    worksheet.append(header)

    # --- Generate and append data rows ---
    for i in range(rows):
        row_data = [f"R{i+1}C{j+1}" for j in range(cols - 1)]
        # Add a random number in the last column for variety
        row_data.append(random.randint(1, 10000))
        worksheet.append(row_data)
        
        # Optional: Print progress for very large files
        if (i + 1) % 10000 == 0:
            print(f"  ... {i+1}/{rows} rows written.")

    # --- Save the workbook to a file ---
    print(f"\nSaving workbook to '{filename}'...")
    workbook.save(filename)
    print("Workbook saved successfully.")


if __name__ == "__main__":   
    # --- Performance Measurement Setup ---
    start_mem = get_memory_usage_mb()
    start_time = time.perf_counter()

    print("--- Performance Test Start ---")
    print(f"Initial Memory Usage: {start_mem:.2f} MB")
    print("-" * 30)

    # --- Main Task ---
    # Define the size of the spreadsheet
    NUM_ROWS = 100000
    NUM_COLS = 50
    OUTPUT_FILENAME = "performance_test_file.xlsx"
    Path(OUTPUT_FILENAME).unlink(True)

    create_large_xlsx(filename=OUTPUT_FILENAME, rows=NUM_ROWS, cols=NUM_COLS, write_only=True)

    # --- Performance Measurement End ---
    end_time = time.perf_counter()
    end_mem = get_memory_usage_mb()

    elapsed_time = end_time - start_time
    peak_mem_usage = end_mem - start_mem
    file_size = os.path.getsize(OUTPUT_FILENAME) / (1024 * 1024)

    print("-" * 30)
    print("--- Performance Test Results ---")
    print(f"Process finished in: {elapsed_time:.4f} seconds")
    print(f"Peak memory usage during process: {peak_mem_usage:.2f} MB")
    print(f"Final memory usage: {end_mem:.2f} MB")
    print(f"Final file size: {file_size:.2f} MB")
    print("-" * 30)
