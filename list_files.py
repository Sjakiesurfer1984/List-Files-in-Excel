import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
import os

# To use this script, you must first install the openpyxl library:
# In your terminal, run: pip install openpyxl

def select_files_and_create_spreadsheet():
    """
    Opens a file selection dialog, gets the names of the selected files,
    and saves them to a new Excel spreadsheet (.xlsx).
    """
    # Create a hidden Tkinter root window
    root = tk.Tk()
    root.withdraw()

    # Open a file dialog to allow the user to select multiple files.
    # The initialdir can be set to a specific path if desired.
    print("Please select the files you want to list in the spreadsheet.")
    selected_files = filedialog.askopenfilenames(
        title="Select Files",
        filetypes=[("All files", "*.*")]
    )

    # Check if the user selected any files
    if not selected_files:
        print("No files were selected. Operation canceled.")
        return

    # Prompt the user for a save location and filename for the spreadsheet
    print("\nPlease choose a location and filename to save the spreadsheet.")
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Spreadsheet As"
    )

    # Check if the user provided a save path
    if not save_path:
        print("Save location not specified. Spreadsheet not created.")
        return

    try:
        # Create a new Excel workbook and get the active worksheet
        workbook = Workbook()
        sheet = workbook.active
        
        # Set a header for the first column
        sheet['A1'] = "File Name (with extension)"

        # Loop through the selected files and add their base names to column A
        for index, file_path in enumerate(selected_files, start=2):
            # Get just the filename from the full path
            file_name = os.path.basename(file_path)
            # Write the filename to the current row in column A
            sheet[f'A{index}'] = file_name

        # Save the workbook to the specified path
        workbook.save(save_path)
        print(f"\nSuccess! Spreadsheet saved to: {save_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

# The standard entry point for running the script
if __name__ == "__main__":
    select_files_and_create_spreadsheet()

