import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import xml.etree.ElementTree as ET
import json
import os  # To handle file paths

def get_columns_from_csv(file_path):
    with open(file_path, newline='') as csvfile:
        reader = csv.reader(csvfile)
        columns = next(reader)  # Read the header row to get the column names
        return columns

def get_columns_from_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    namespaces = {
        'saw': 'com.siebel.analytics.web/report/v1.1',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'sawx': 'com.siebel.analytics.web/expression/v1.1'
    }

    expr_elements = root.findall(".//saw:column//sawx:expr[@xsi:type='sawx:sqlExpression']", namespaces=namespaces)
    columns = [elem.text.replace('"', '').lower() for elem in expr_elements]
    return columns

def get_columns_from_json(file_path):
    with open(file_path) as jsonfile:
        data = json.load(jsonfile)
        columns = list(data[0].keys()) if data else []
        return columns

def get_columns_from_file(file_path):
    if file_path.lower().endswith('.csv'):
        return get_columns_from_csv(file_path)
    elif file_path.lower().endswith('.xml'):
        return get_columns_from_xml(file_path)
    elif file_path.lower().endswith('.json'):
        return get_columns_from_json(file_path)
    else:
        raise ValueError("Unsupported file format")

def append_to_csv(all_columns, output_file, message=""):
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        if message:
            writer.writerow([message])  # Write message if no merge happens
        writer.writerow(all_columns)  # Write column headers
    print(f"Combined columns have been written to {output_file}")

def calculate_matching_percentage(columns_files):
    # Get the intersection and union of columns for multiple files
    all_columns = [set(columns) for columns in columns_files]
    matching_columns = set.intersection(*all_columns)
    total_columns = set.union(*all_columns)
    
    if total_columns == 0:
        return 0
    
    matching_percentage = (len(matching_columns) / len(total_columns)) * 100
    return matching_percentage

def compare_and_combine_columns(file_paths):
    columns_files = [get_columns_from_file(file_path) for file_path in file_paths]
    matching_percentage = calculate_matching_percentage(columns_files)
    print(f"Matching Percentage: {matching_percentage:.2f}%")

    file_name_combined = "_".join([os.path.splitext(os.path.basename(file))[0] for file in file_paths])
    output_file_path = f"{file_name_combined}_merged.csv"

    if matching_percentage > 50:
        # Combine and deduplicate columns across all files
        all_columns = sorted(set().union(*columns_files))
        append_to_csv(all_columns, output_file_path)
    else:
        append_to_csv([], output_file_path, message=f"Matching percentage is {matching_percentage:.2f}%. Not merging.")
    
    return output_file_path, matching_percentage

# GUI for file selection and output display
class ColumnComparerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Report Comparison Tool")

        # Labels, Buttons, and Text boxes
        self.label1 = tk.Label(root, text="Select files:")
        self.label1.grid(row=0, column=0, padx=10, pady=10)

        self.file_button = tk.Button(root, text="Browse", command=self.select_files)
        self.file_button.grid(row=0, column=1, padx=10, pady=10)

        # Label to display the selected files
        self.file_label = tk.Label(root, text="No files selected", anchor="w")
        self.file_label.grid(row=0, column=2, padx=10, pady=10, sticky="ew")

        self.compare_button = tk.Button(root, text="Compare", state=tk.DISABLED, command=self.compare_columns)
        self.compare_button.grid(row=1, column=0, columnspan=2, pady=20)

        # Result label
        self.result_label = tk.Label(root, text="Results will be shown here", anchor="w", width=50, height=4)
        self.result_label.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        # Initialize file paths
        self.files = []

        # Make sure the window resizes appropriately
        root.grid_rowconfigure(0, weight=0)  # Row for the file selection buttons
        root.grid_rowconfigure(1, weight=0)  # Row for the compare button
        root.grid_rowconfigure(2, weight=1)  # Row for results (this row will expand)
        root.grid_columnconfigure(2, weight=1)  # Make file label and result label expand horizontally

    def select_files(self):
        self.files = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv"), ("XML Files", "*.xml"), ("JSON Files", "*.json")])
        if self.files:
            # Update the file name label to show the selected files
            self.file_label.config(text=", ".join([os.path.basename(file) for file in self.files]))
        if self.files:
            self.compare_button.config(state=tk.NORMAL)

    def compare_columns(self):
        if not self.files:
            messagebox.showerror("Error", "Please select files!")
            return
        
        try:
            output_file_path, matching_percentage = compare_and_combine_columns(self.files)
            result_message = f"Matching Percentage: {matching_percentage:.2f}%\n"
            result_message += f"Output File: {output_file_path}"
            self.result_label.config(text=result_message)
            messagebox.showinfo("Comparison Complete", f"Comparison complete. Results saved to {output_file_path}.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Running the Tkinter app
if __name__ == "__main__":
    root = tk.Tk()
    app = ColumnComparerApp(root)
    root.mainloop()

