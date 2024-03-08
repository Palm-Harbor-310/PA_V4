import os
import pandas as pd
import re

folder_path = "P:/Temp"

def pullPricesFromExcel(folder_path):
    """
    Method to pull prices from Excel files in a folder and process them.

    Args:
    folder_path (str): Path to the folder containing Excel files.
    """

    # Get a list of all Excel files in the folder
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    for file in excel_files:
        # Construct full path to the current file
        file_path = os.path.join(folder_path, file)

        # Read the current Excel file
        original_df = pd.read_excel(file_path)

        # Identify the column with Part Numbers using regex
        part_number_column = None
        part_number_pattern = re.compile(r'P\d{2}-\d{3}-\d{3}')

        for column in original_df.columns:
            if original_df[column].astype(str).str.match(part_number_pattern).any():
                part_number_column = column
                break

        if part_number_column is None:
            raise ValueError("Part number column not found.")
        else:
            original_df.rename(columns={part_number_column: 'pr_codenum'}, inplace=True)
        
        # Load the LPC Excel file
        lpc_df = pd.read_excel(
            r"C:\Users\daniel.pace\Documents\Coding\Purchasing Automation\PA_V4\Price Comparison\Last Part Cost 3.1.xlsx",
            sheet_name='All')
        
        # Map PO Cost from LPC to original dataframe based on part numbers
        original_df = original_df.merge(lpc_df[['pr_codenum', 'PO Cost']], on='pr_codenum', how='left')

        # Save the processed data to a new file with the original filename
        original_df.to_excel(os.path.join(folder_path, f"processed_{file}"), index=False)

# Example usage
# Replace with your actual folder path
pullPricesFromExcel(folder_path)
