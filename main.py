import os
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path


def combine_csv_to_excel(csv_folder, output_excel_file):
    # List to hold dataframes
    dataframes = []

    # Iterate over all CSV files in the folder
    for file in os.listdir(csv_folder):
        if file.endswith('.csv'):
            file_path = os.path.join(csv_folder, file)
            # Read CSV into a dataframe
            df = pd.read_csv(file_path, header=None)

            df.insert(0, 'Source.Name', [file] * (len(df)))
            dataframes.append(df)

    # Check if there are any CSV files
    if not dataframes:
        msg = "No CSV files found in the folder."
        print(msg)
        return False, msg

    # Concatenate all dataframes
    combined_df = pd.concat(dataframes, ignore_index=True)

    # Save to Excel
    combined_df.to_excel(output_excel_file, index=False, engine='openpyxl')

    print(f"Combined data has been saved to {output_excel_file}")

    return True, "Done. Saved to "+output_excel_file


def split_excel_to_csv(input_excel_file, output_folder):
    # Read the Excel file into a dataframe
    df = pd.read_excel(input_excel_file, engine='openpyxl')

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    if 'Source.Name' not in df.columns:
        msg = "The 'Source.Name' column is missing in the Excel file."
        print(msg)
        return False, msg

    # Group by the 'Source.Name' column and save each group as a CSV file
    for source_file, group in df.groupby('Source.Name'):
        # Remove the 'Source.Name' column before saving
        group = group.drop(columns=['Source.Name'])

        # Save the group to a CSV file
        output_csv_file = os.path.join(output_folder, source_file)
        group.to_csv(output_csv_file, index=False, header=False)

        print(f"Saved data for '{source_file}' to {output_csv_file}")

    return True, "Done. Saved to "+output_folder


class FileInputApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to XLSX")
        # self.root.geometry("600x200")

        # Create and configure main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Folder input section
        self.folder_frame = ttk.LabelFrame(
            self.main_frame, text="Append multiple CSV to Excel", padding="5")
        self.folder_frame.grid(
            row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.folder_path = tk.StringVar()
        self.folder_entry = ttk.Entry(
            self.folder_frame, textvariable=self.folder_path, width=50)
        self.folder_entry.grid(row=0, column=0, padx=5)

        self.folder_button = ttk.Button(
            self.folder_frame, text="Select folder", command=self.browse_folder)
        self.folder_button.grid(row=0, column=1, padx=5)

        self.process_button = ttk.Button(
            self.folder_frame, text="Process", command=self.save_xlsx)
        self.process_button.grid(row=0, column=2, padx=5)

        # File input section
        self.file_frame = ttk.LabelFrame(
            self.main_frame, text="Split Excel to multiple CSV", padding="5")
        self.file_frame.grid(row=1, column=0, columnspan=2,
                             sticky=(tk.W, tk.E), pady=5)

        self.file_path = tk.StringVar()
        self.file_entry = ttk.Entry(
            self.file_frame, textvariable=self.file_path, width=50)
        self.file_entry.grid(row=0, column=0, padx=5)

        self.file_button = ttk.Button(
            self.file_frame, text="Select file", command=self.browse_file)
        self.file_button.grid(row=0, column=1, padx=5)

        self.process_file_button = ttk.Button(
            self.file_frame, text="Process", command=self.split_xlsx)
        self.process_file_button.grid(row=0, column=2, padx=5)

        # Status label
        self.status_label = ttk.Label(self.main_frame, text="")
        self.status_label.grid(row=2, column=0, columnspan=2, pady=10)

    def save_xlsx(self):
        if not self.folder_path.get():
            tk.messagebox.showwarning(
                title="Warning", message="Please select the CSV folder first")
            return

        postfix = time.strftime(".%d-%m-%Y %H.%M.xlsx")
        output_path = self.folder_path.get()+postfix

        print(output_path)

        success, msg = combine_csv_to_excel(
            self.folder_path.get(), output_path)

        self.status_label.config(
            text=msg, foreground="green" if success else "red")

    def split_xlsx(self):
        if not self.file_path.get():
            tk.messagebox.showwarning(
                title="Warning", message="Please select the Excel file first")
            return

        file_path = Path(self.file_path.get())
        file_name = file_path.stem.split('.')[0]

        postfix = time.strftime("-split %d-%m-%Y %H.%M")
        output_path = str(Path(file_path.parent, file_name+postfix))

        print(output_path)

        success, msg = split_excel_to_csv(
            self.file_path.get(), output_path)

        self.status_label.config(
            text=msg, foreground="green" if success else "red")

    def browse_folder(self):
        folder_selected = filedialog.askdirectory(
            title="Select Folder containing all CSV files")
        if folder_selected:
            self.folder_path.set(folder_selected)

    def browse_file(self):
        file_selected = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"),
                       ("Excel files", "*.xls"),
                       ("Excel files", "*.xlsm")]
        )
        if file_selected:
            self.file_path.set(file_selected)


def main():
    root = tk.Tk()
    app = FileInputApp(root)
    root.mainloop()


main()
