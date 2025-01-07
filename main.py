import os
import pandas as pd

def combine_csv_to_excel(csv_folder, output_excel_file):
    # List to hold dataframes
    dataframes = []

    # Iterate over all CSV files in the folder
    for file in os.listdir(csv_folder):
        if file.endswith('.csv'):
            file_path = os.path.join(csv_folder, file)
            # Read CSV into a dataframe
            df = pd.read_csv(file_path)
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

# Folder containing the CSV files
csv_folder = "./csv"

# Output Excel file
output_excel_file = "combined_data.xlsx"

# Combine CSV files into an Excel file
combine_csv_to_excel(csv_folder, output_excel_file)

# csv_file = "/home/bagas/Downloads/csv/awooga.csv"

# df = pd.read_csv(csv_file)
# xlsx_file = os.path.splitext(csv_file)[0] + '.xlsx'
# df.to_excel(xlsx_file, index=None, header=True)
