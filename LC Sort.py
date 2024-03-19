
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

def show_instructions_and_confirm():
    # Create a root window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Show an instruction message box with OK and Cancel options
    response = messagebox.askokcancel("Instructions", "This application is designed to function with .xlsx files that have Library of Congress call numbers in column A, with the first row serving as the header row. Column A must be labeld 'Call Number'.\n\nIn the following window, you will be prompted to choose your input file. The output file will be saved in the same folder as the input file and will have '_LCSortable' appended to the filename. The last column of the sheet will be the 'Sortable Call Number' column.\n\nPress 'OK' to proceed or 'Cancel' to exit.")
    
    root.destroy()
    return response

def select_input_file():
    root = tk.Tk()
    root.withdraw()  # We don't want a full GUI, so keep the root window from appearing
    while True:
        filename = filedialog.askopenfilename()  # Show an "Open" dialog box and return the path to the selected file
        if not filename:
            messagebox.showwarning("No file selected", "No file was selected. The application will now exit.")
            root.destroy()
            sys.exit()
        if filename.endswith('.xlsx'):
            return filename
        messagebox.showwarning("Incorrect file type", "Please select an XLSX file.")

def alert_and_exit(message):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showerror("Error", message)
    root.destroy()
    sys.exit()

# Call the function to show instructions and wait for user confirmation
if show_instructions_and_confirm():
    # User pressed OK, proceed with selecting the input file
    input_file_path = select_input_file()
else:
    # User pressed Cancel, exit the script
    sys.exit()

import pandas as pd
import os
import pycallnumber as pycn

if input_file_path:
    # Reading the input Excel file into a DataFrame
    df = pd.read_excel(input_file_path)

    # Ensure call numbers are treated as strings
    df['Call Number'] = df['Call Number'].astype(str)

    # Check if the first column is labeled 'Call Number'
    if df.columns[0] != 'Call Number':
        alert_and_exit("The call numbers must be in the first column, labeled as 'Call Number'. Please correct the file and try again.")

    # Initialize a list to store processed rows
    all_rows = []

    # Iterate through DataFrame rows
    for index, row in df.iterrows():
        call_number_str = row['Call Number']  # Assuming the call number is in the first column
        # Check if the call number is 'N/A' or a similar placeholder
        if call_number_str in ['N/A', 'NaN', 'nan', '']:
            # Assign a placeholder value for sorting or bypass sorting logic
            sortable_call_number = 'ZZZZ'  # Placeholder to ensure these go at the end; adjust as needed
        else:
            try:
                call_number = pycn.callnumber(call_number_str)
                sortable_call_number = call_number.for_sort()
            except Exception as e:
                print(f"Error processing call number '{call_number_str}': {e}")
                sortable_call_number = 'ZZZZ'  # Handle error in sorting call number

        # Add the sortable call number to the row
        row['Sortable Call Number'] = sortable_call_number
        all_rows.append(row)

    # Convert the list of rows back into a DataFrame
    df_processed = pd.DataFrame(all_rows)

    # Optionally, sort the DataFrame by the 'Sortable Call Number' column if desired
    # df_processed = df_processed.sort_values(by=['Sortable Call Number'])

    # Constructing the output filename
    filename_without_ext, file_ext = os.path.splitext(input_file_path)
    output_file_path = filename_without_ext + "_LCSortable.xlsx"

    # Save the processed DataFrame to an Excel file
    df_processed.to_excel(output_file_path, index=False)

# Display a completion message
root = tk.Tk()
root.withdraw()  # Hide the main window
messagebox.showinfo("Process Completed", "The processing is complete. The output file has been saved.")
root.destroy()