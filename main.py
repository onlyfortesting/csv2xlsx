import os
import pandas as pd
from pprint import pprint


def combine_csv_to_excel(csv_folder, output_excel_file):
    # List to hold dataframes
    dataframes = []

    # Iterate over all CSV files in the folder
    for file in os.listdir(csv_folder):
        if file.endswith('.csv'):
            file_path = os.path.join(csv_folder, file)
            # Read CSV into a dataframe
            df = pd.read_csv(file_path)
            # df.insert(0, 'Source File', [file] + [None] * (len(df) - 1))
            df.insert(0, 'Source File', [file] * (len(df)))
            dataframes.append(df)

    # Check if there are any CSV files
    if not dataframes:
        print("No CSV files found in the folder.")
        return

    # Concatenate all dataframes
    combined_df = pd.concat(dataframes, ignore_index=True)

    # Save to Excel
    # combined_df.to_excel(output_excel_file, index=False, engine='openpyxl')
    combined_df.to_excel(output_excel_file, index=False)

    print(f"Combined data has been saved to {output_excel_file}")


def split_excel_to_csv(input_excel_file, output_folder):
    # Read the Excel file into a dataframe
    df = pd.read_excel(input_excel_file, engine='openpyxl')

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Check if the 'Source File' column exists
    if 'Source File' not in df.columns:
        print("The 'Source File' column is missing in the Excel file.")
        return

    # Group by the 'Source File' column and save each group as a CSV file
    for source_file, group in df.groupby('Source File'):
        # Remove the 'Source File' column before saving
        group = group.drop(columns=['Source File'])
        # print(group)

        # Save the group to a CSV file
        output_csv_file = os.path.join(output_folder, source_file)
        group.to_csv(output_csv_file, index=False)

        print(f"Saved data for '{source_file}' to {output_csv_file}")


# Folder containing the CSV files
csv_folder = "./csv"

# Output Excel file
output_excel_file = "combined_data.xlsx"

# Combine CSV files into an Excel file
# combine_csv_to_excel(csv_folder, output_excel_file)
split_excel_to_csv("combined_data.xlsx", "splitted")
