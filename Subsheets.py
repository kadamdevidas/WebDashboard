#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct 10 19:48:11 2024

@author: omkarkadam
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Oct 10 12:17:49 2024

@author: kadam.omkar
"""

import pandas as pd
import os

# Folder containing the monthly MPR Excel files
folder_path = r"/Users/omkarkadam/Documents/Automate with Python/Dashboard" # Replace with the path to your folder

# List of months for which MPR files are present
months = ['JUNE', 'JULY']  # Modify this based on the actual files in the folder

# Function to extract tables from a single Excel file
def extract_tables_from_file(file_path):
    # Read the first sheet of the Excel file and skip unnecessary rows
    df = pd.read_excel(file_path, sheet_name=0, skiprows=5)
    
    # Extract the tables
    table_1 = df.iloc[1:4, 1:4]
    table_1.columns = ['Plant', 'Planned', 'Executed']
    
    table_2 = df.iloc[7:10, 1:4]
    table_2.columns = ['Plant', 'Planned', 'Executed']
    
    table_3 = df.iloc[13:16, 1:3]
    table_3.columns = ['Plant', 'No. of Breakdown Jobs']
    
    return table_1, table_2, table_3

# Initialize dictionaries to store tables for each month
merged_tables = {
    "table_1": [],
    "table_2": [],
    "table_3": []
}

# Loop through each month, read the corresponding file, and extract tables
for month in months:
    file_name = f'MPR {month}.xlsx'  # Construct the filename
    file_path = os.path.join(folder_path, file_name)
    
    # Extract tables from the current month's file
    table_1, table_2, table_3 = extract_tables_from_file(file_path)
    
    # Add a 'Month' column to each table
    table_1['Month'] = month
    table_2['Month'] = month
    table_3['Month'] = month
    
    # Append tables to the respective list in merged_tables
    merged_tables["table_1"].append(table_1)
    merged_tables["table_2"].append(table_2)
    merged_tables["table_3"].append(table_3)

# Concatenate tables from all months
merged_table_1 = pd.concat(merged_tables["table_1"], ignore_index=True)
merged_table_2 = pd.concat(merged_tables["table_2"], ignore_index=True)
merged_table_3 = pd.concat(merged_tables["table_3"], ignore_index=True)

# Create a new Excel writer to save the merged tables
output_file = r"/Users/omkarkadam/Documents/Automate with Python/Merged_MPR.xlsx" # Replace with your desired output path

# Write each merged table to its own sheet using the new sheet names
with pd.ExcelWriter(output_file) as writer:
    # Write each merged table to its respective sheet
    merged_table_1.to_excel(writer, sheet_name='PM02', index=False)
    merged_table_2.to_excel(writer, sheet_name='PM01', index=False)
    merged_table_3.to_excel(writer, sheet_name='Breakdown Maintenance', index=False)

print(f"Merged file created successfully with new sheet names at {output_file}")
