import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Function to count occurrences of values in 'Direction' column
def count_directions(df):
    anterograde = (df['Direction'] == 1).sum()
    retrograde = (df['Direction'] == -1).sum()
    stationary = (df['Direction'] == 0).sum()
    total = len(df)
    return anterograde, retrograde, stationary, total

# Function to calculate percentages
def calculate_percentages(antro, retro, stationary, total):
    percent_antro = (antro / total) * 100
    percent_retro = (retro / total) * 100
    percent_stationary = (stationary / total) * 100
    return percent_antro, percent_retro, percent_stationary

# Initialize tkinter
root = tk.Tk()
root.withdraw()

# Ask user to select a directory
directory = filedialog.askdirectory(title='Select Input Directory')

# Check if a directory was selected
if not input_directory:
    print("No directory selected. Exiting...")
    exit()

# Initialize empty DataFrame to store results
result_df = pd.DataFrame(columns=['Direction'])

# Iterate over files in the directory
for filename in os.listdir(directory):
    if filename.startswith('kymograph') and (filename.endswith('.xlsx') or filename.endswith('.xls')):
        filepath = os.path.join(directory, filename)
        # Read Excel file into DataFrame
        df = pd.read_excel(filepath)
        # Count occurrences of directions
        antro, retro, stationary, total = count_directions(df)
        # Calculate percentages
        percent_antro, percent_retro, percent_stationary = calculate_percentages(antro, retro, stationary, total)
        # Add results to DataFrame
        result_df[filename] = [antro, retro, stationary, total, percent_antro, percent_retro, percent_stationary]

# Save DataFrame to Excel file
result_df.to_excel(os.path.join(directory, 'directionresults.xlsx'), index=False)

