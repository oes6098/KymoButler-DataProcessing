import os
import pandas as pd
import numpy as np
from tkinter import filedialog
import tkinter as tk

# Create a Tkinter root window
root = tk.Tk()
root.withdraw()  # Hide the root window

# Ask user to select the input directory using a file dialog
input_directory = filedialog.askdirectory(title="Select Input Directory")

# Check if a directory was selected
if not input_directory:
    print("No directory selected. Exiting...")
    exit()

# Concatenate the directory path with the new folder name
results_folder_path = os.path.join(input_directory, "kymoresults")

# Create the new folder
os.makedirs(results_folder_path, exist_ok=True)


def process_data (direction, column_title):

    # Create a list to store NumPy arrays
    result_arrays = []

    # Iterate over files in the directory
    for file_name in os.listdir(input_directory):
        file_path = os.path.join(input_directory, file_name)
        
        # Check if the item in the directory is an Excel file
        if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            print(f'Processing file: {file_name}')
            
            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(file_path)

            # Convert 'track total displacement [um]' and 'Direction' to numeric values
            df[column_title] = pd.to_numeric(df[column_title], errors='coerce')
            df['Direction'] = pd.to_numeric(df['Direction'], errors='coerce')

            if direction == 'ANT':
                # Filter rows starting from the 2nd row based on conditions
                filtered_rows = df.iloc[0:].loc[(~df[column_title].isna()) & (df['Direction'] == 1)]
            elif direction == 'RET':
                filtered_rows = df.iloc[0:].loc[(~df[column_title].isna()) & (df['Direction'] == -1)]
            elif direction == 'STAT':
                filtered_rows = df.iloc[0:].loc[(~df[column_title].isna()) & (df['Direction'] == 0)]
            else:
                filtered_rows = df.iloc[0:].loc[(~df[column_title].isna())]

            # Extract values from column 'track total displacement [um]' and convert them to a NumPy array
            result_array = np.array(filtered_rows[column_title])

            # Add the NumPy array to the list
            result_arrays.append(result_array)

    if result_arrays:
        # Create a DataFrame from the list of arrays
        result_df = pd.DataFrame(result_arrays).T
        result_df.columns = [f'Column_{i+1}' for i in range(result_df.shape[1])]

        if 'frame2frame' in column_title:
            # Save the resulting DataFrame to a CSV file
            result_df.to_csv(os.path.join(input_directory, 'kymoresults', f"{direction}_frame2framevelocity.csv"), index=False)
            print("Results saved successfully.")
        elif 'Start2end' in column_title: 
            result_df.to_csv(os.path.join(input_directory, 'kymoresults', f"{direction}_start2endvelocity.csv"), index=False)
            print("Results saved successfully.")
        else:
            result_df.to_csv(os.path.join(input_directory, 'kymoresults', direction+column_title+'.csv'), index=False)
            print("Results saved successfully.")
    else:
        print("No valid data found to save.")

process_data( 'ANT', 'Av frame2frame velocity [um/sec]')
process_data('RET', 'Av frame2frame velocity [um/sec]')
process_data('STAT', 'Av frame2frame velocity [um/sec]')
process_data('TOTAL', 'Av frame2frame velocity [um/sec]')
process_data('ANT', 'Start2end velocity [um/sec]')
process_data('RET', 'Start2end velocity [um/sec]')
process_data('STAT', 'Start2end velocity [um/sec]')
process_data('TOTAL', 'Start2end velocity [um/sec]')
process_data('ANT', 'track duration [sec]')
process_data('RET', 'track duration [sec]')
process_data('STAT', 'track duration [sec]')
process_data('TOTAL', 'track duration [sec]')
process_data('ANT', 'track total displacement [um]')
process_data('RET', 'track total displacement [um]')
process_data('STAT', 'track total displacement [um]')
process_data('TOTAL', 'track total displacement [um]')


#Direction analysis

# Function to count occurrences of values in 'Direction' column
def count_directions(df):
    anterograde = (df['Direction'] == 1).sum()
    retrograde = (df['Direction'] == -1).sum()
    stationary = (df['Direction'] == 0).sum()
    total = len(df)
    return anterograde, retrograde, stationary, total

# Function to calculate percentages
def calculate_percentages(antero, retro, stationary, total):
    percent_antro = (antero / total) * 100
    percent_retro = (retro / total) * 100
    percent_stationary = (stationary / total) * 100
    return percent_antro, percent_retro, percent_stationary

# Initialize empty DataFrame to store results
result_df = pd.DataFrame(columns=[' '])

# Define the values for the first column
first_column_values = [ 'anterograde', 'retrograde', 'stationary', 'total', 'percent anterograde', 'percent retrograde', 'percent stationary']

# Insert the values into the first column of the dataframe
result_df.insert(0, 'Direction', first_column_values)

print(result_df)

#Iterate over files in the directory
for filename in os.listdir(input_directory):
    if filename.startswith('kymograph') and (filename.endswith('.xlsx') or filename.endswith('.xls')):
        filepath = os.path.join(input_directory, filename)
        # Read Excel file into DataFrame
        df = pd.read_excel(filepath)
        # Count occurrences of directions
        antero, retro, stationary, total = count_directions(df)
        # Calculate percentages
        percent_antro, percent_retro, percent_stationary = calculate_percentages(antero, retro, stationary, total)
        # Add results to DataFrame
        result_df[filename] = [antero, retro, stationary, total, percent_antro, percent_retro, percent_stationary]

# Save DataFrame to Excel file
result_df.to_csv(os.path.join(input_directory, 'kymoresults', 'directionresults.csv'), index=False)