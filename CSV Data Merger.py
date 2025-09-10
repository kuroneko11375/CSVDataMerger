"""
CSV Data Merger

一個用於將多個 CSV 檔案中特定欄位合併到單一 Excel 檔案的 Python 腳本。
這個工具專門用來處理符合特定命名模式的 CSV 檔案，並從預定的資料列範圍中擷取資料。

Author:SchwarzeKatze_R
License: MIT
Version: 1.0.0
"""

import os
import glob
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime
from typing import Optional


def pick_folder() -> str:
    """
    Display a folder selection dialog for the user.
    
    Returns:
        str: The selected folder path, or empty string if cancelled.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory(title="Select folder containing CSV files")
    return folder_path


def merge_csv_files(folder_path: str, file_prefix: str = "DATA", column_index: int = 3, 
                   start_row: int = 55, num_rows: int = 12) -> None:
    """
    Merge data from multiple CSV files with specific naming pattern.
    
    This function searches for CSV files named in the pattern:
    {file_prefix}1_RESULT_YYYYMMDD_*.csv to {file_prefix}6_RESULT_YYYYMMDD_*.csv
    
    Args:
        folder_path (str): Path to the folder containing CSV files.
        file_prefix (str): Prefix for CSV file names (default: "DATA").
        column_index (int): Column index to extract (0-based, default: 3 for column D).
        start_row (int): Starting row number (1-based, default: 55).
        num_rows (int): Number of rows to extract (default: 12).
    """
    date_str = datetime.now().strftime("%Y%m%d")
    
    # Store data series from each file
    all_series = []
    
    for i in range(1, 7):
        # Pattern for CSV files
        pattern = os.path.join(folder_path, f"{file_prefix}{i}_RESULT_{date_str}_*.csv")
        files = glob.glob(pattern)
        
        if files:
            # Use the first matching file
            csv_file = files[0]
            print(f"Processing: {os.path.basename(csv_file)}")
            
            try:
                df_temp = pd.read_csv(
                    csv_file,
                    skiprows=start_row-1,  # Convert to 0-based indexing
                    nrows=num_rows,
                    header=None,
                    usecols=[column_index],
                    engine="python"
                )
                s_temp = df_temp.iloc[:, 0]
                s_temp.name = f"{file_prefix}{i}"
                all_series.append(s_temp)
            except Exception as e:
                print(f"Error reading file {csv_file}: {e}")
                # Add empty series if file cannot be read
                s_temp = pd.Series([None]*num_rows, name=f"{file_prefix}{i}")
                all_series.append(s_temp)
        else:
            print(f"No file found for pattern: {file_prefix}{i}_RESULT_{date_str}_*.csv")
            # Add empty series if no file found
            s_temp = pd.Series([None]*num_rows, name=f"{file_prefix}{i}")
            all_series.append(s_temp)
            
    # Merge all series horizontally
    df_final = pd.concat(all_series, axis=1)
    
    # Output to Excel file
    output_file = os.path.join(folder_path, f"merged_data_{date_str}.xlsx")
    try:
        df_final.to_excel(output_file, index=False)
        print(f"Data merged successfully. Output file: {output_file}")
        print(f"Final data shape: {df_final.shape}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")


def main() -> None:
    """
    Main entry point of the application.
    
    Prompts user to select a folder and processes CSV files within it.
    """
    print("CSV Data Merger")
    print("=" * 50)
    
    # Let user select folder via GUI
    folder_selected = pick_folder()
    
    # Exit if no folder selected
    if not folder_selected:
        print("No folder selected. Exiting program.")
        return
    
    print(f"Selected folder: {folder_selected}")
    
    # You can customize these parameters as needed:
    # - file_prefix: Change this to match your file naming pattern
    # - column_index: Column to extract (0=A, 1=B, 2=C, 3=D, etc.)
    # - start_row: Starting row number (1-based)
    # - num_rows: Number of rows to extract
    
    merge_csv_files(
        folder_path=folder_selected,
        file_prefix="DATA",  # Change this to your actual file prefix
        column_index=3,      # Column D (0-based index)
        start_row=55,        # Start from row 55
        num_rows=12          # Extract 12 rows
    )


if __name__ == "__main__":
    main()
