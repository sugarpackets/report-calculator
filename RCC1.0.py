#Risk and Compliance Calculator, developed by Brandon Price, Safety and Infection Control Coordinator, Brevard Health Alliance, 2024.
#Users are free to recreate, modify, and distribute this software unless my company eventually tells me it's theirs idk/idc.
#Users may need to ask IT to install python and libraries for this to work:
#pip install pandas openpyxl
#Intent is to use a results .xlsx from Google Forms and tally the clinic name (column H) and date (column I) to count 
#the amount of times a clinic has performed an observation
#then export this data in a table that meets reporting criteria.
#Import pandas is the module to analyze data tables like xlsx.
#Import os is the module for grabbing directory names to be used in file naming and export placement.
#From datetime import datetime is the module for adding a timestamp to the export file.
#Import platform determines the OS so the file knows HOW to open the result export.
#Import subprocess allows the execution of opening the export file.
#Import tkinter module required for app UI so my boss doesn't scream from seeing a CLI.
#Hardcoded requirements are determined by Risk and Compliance using clinic staff popuation.
#This application has no inherent reliance on internet connectivity and should pass no data through the network UNLESS
#the INPUT or OUTPUT file locations are not locally stored.

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
import platform
import subprocess

# Hardcoded requirements
requirements = {
    "Barton": 10,
    "Circles of Care": 2,
    "Endeavor": 5,
    "Evans": 2,
    "Heritage Park": 2,
    "Malabar": 5,
    "Mobile": 2,
    "Palm Bay": 10,
    "Port St John": 5,
    "Sarno": 10,
    "Valentine": 5,
    "Titusville Dental": 2,
    "Titusville": 5,
    "University Dental": 2
}

def get_fiscal_quarter(date):
    """Function to calculate the fiscal quarter, starts in October."""
    month = date.month
    if month in [10, 11, 12]:  # Q1 for fiscal year
        return "Q1"
    elif month in [1, 2, 3]:   # Q2 for fiscal year
        return "Q2"
    elif month in [4, 5, 6]:   # Q3 for fiscal year
        return "Q3"
    else:                      # Q4 for fiscal year
        return "Q4"

def tally_and_compare(file_path):
    # Load columns H (clinic location) and I (date of occurrence)
    df = pd.read_excel(file_path, usecols=[7, 8], names=['Location', 'Date'])

    # Convert 'Date' to datetime
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    # Drop rows where 'Date' is NaT (not a valid datetime)
    df = df.dropna(subset=['Date'])

    # Add a new column 'Fiscal_Quarter' by applying the fiscal quarter calculation
    df['Fiscal_Quarter'] = df['Date'].apply(get_fiscal_quarter)

    # Get all possible combinations of locations and fiscal quarters
    quarters = ["Q1", "Q2", "Q3", "Q4"]
    full_index = pd.MultiIndex.from_product([requirements.keys(), quarters], names=["Location", "Fiscal_Quarter"])

    # Group by 'Location' and 'Fiscal_Quarter' and count the occurrences
    grouped = df.groupby(['Location', 'Fiscal_Quarter']).size().reindex(full_index, fill_value=0).reset_index(name='Tally')

    # Create a DataFrame for holding percentages
    percentage_df = pd.DataFrame(columns=["Location", "Fiscal_Quarter", "Percentage Met"])

    for _, row in grouped.iterrows():
        location = row['Location']
        fiscal_quarter = row['Fiscal_Quarter']
        tally = row['Tally']

        if location in requirements:
            requirement = requirements[location]
            percentage_met = (tally / requirement) * 100 if requirement > 0 else 0
            # Cap the percentage at 100%
            percentage_met = min(percentage_met, 100)
            # Add the result to the percentage DataFrame
            percentage_df = pd.concat([percentage_df, pd.DataFrame({
                "Location": [location],
                "Fiscal_Quarter": [fiscal_quarter],
                "Percentage Met": [percentage_met]
            })], ignore_index=True)

    # Pivot the DataFrame to have Locations as index, Fiscal Quarters as columns, and percentages as values
    pivot_df = percentage_df.pivot(index='Location', columns='Fiscal_Quarter', values='Percentage Met')

    return pivot_df

def export_result(file_path, result_df):
    # Export the pivoted DataFrame to a new Excel file
    result_df.to_excel(file_path, index=True)

def open_file(file_path):
    # Automatically open the output file after saving it
    if platform.system() == 'Windows':
        os.startfile(file_path)
    elif platform.system() == 'Darwin':  # macOS
        subprocess.call(('open', file_path))
    elif platform.system() == 'Linux':
        subprocess.call(('xdg-open', file_path))

def process_file(file_path):
    try:
        # Process the file and generate the result
        result_df = tally_and_compare(file_path)
        
        # Get the Downloads directory dynamically
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        
        # Extract the original file name without extension
        original_file_name = os.path.splitext(os.path.basename(file_path))[0]
        
        # Get the current date and time in the desired format
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        
        # Define the output file name: original name + "RESULTS" + current date/time
        output_file_name = f"{original_file_name}_RESULTS_{current_time}.xlsx"
        output_file_path = os.path.join(downloads_path, output_file_name)
        
        # Export the result
        export_result(output_file_path, result_df)
        messagebox.showinfo("Success", f"Results exported successfully to {output_file_path}")
        
        # Automatically open the file
        open_file(output_file_path)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        process_file(file_path)

def create_ui():
    root = tk.Tk()
    root.title("Risk and Compliance Calculator")
    
    root.configure(bg='dark blue')

    label = tk.Label(root, text="Select the Excel file to process:")
    label.pack(pady=10)

    browse_button = tk.Button(root, text="Browse", command=browse_file)
    browse_button.pack(pady=10)

    root.geometry("500x200")
    root.mainloop()

if __name__ == "__main__":
    create_ui()