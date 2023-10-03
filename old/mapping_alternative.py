import pandas as pd
from fuzzywuzzy import process
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog, ttk

# Sample dataframes (replace this with your own data or loading function)
df1 = pd.DataFrame({'name': ['apple', 'banana', 'cherry']})
df2 = pd.DataFrame({'name': ['aplpe', 'banan', 'grape']})

# Fuzzy matching function
def fuzzy_match(name, choices, threshold=80):
    match, score = process.extractOne(name, choices)
    if score > threshold:
        return match
    return None

def confirm_matches():
    for i, (index, match) in enumerate(matches):
        if checkboxes_vars[i].get() == 1:  # If checkbox is checked
            df1.at[index, 'mapped_name'] = match
    matches_window.destroy()
    messagebox.showinfo("Info", "Matching Completed!")
    print(df1)  # Display results in console

def main_matching():
    global matches_window, matches, checkboxes_vars
    matches = []
    checkboxes_vars = []
    
    for index, row in df1.iterrows():
        matched_name = fuzzy_match(row['name'], df2['name'].tolist())
        if matched_name:
            matches.append((index, matched_name))
    
    matches_window = tk.Toplevel(root)
    matches_window.title("Confirm Matches")
    
    for match in matches:
        var = tk.IntVar()
        ttk.Checkbutton(matches_window, text=f"Map {df1.loc[match[0], 'name']} to {match[1]}", variable=var).pack(anchor="w", padx=10, pady=5)
        checkboxes_vars.append(var)
    
    confirm_button = tk.Button(matches_window, text="Confirm Matches", command=confirm_matches)
    confirm_button.pack(pady=20)

def load_dataframe():
    file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        return pd.read_excel(file_path, engine='openpyxl')
    return None


def load_data():
    global df1, df2
    messagebox.showinfo("Info", "Load the first dataframe.")
    df1 = load_dataframe()
    messagebox.showinfo("Info", "Load the second dataframe.")
    df2 = load_dataframe()

root = tk.Tk()
root.title("Fuzzy Matching GUI")

load_button = tk.Button(root, text="Load Data", command=load_data)
load_button.pack(pady=20)

start_button = tk.Button(root, text="Start Matching", command=main_matching)
start_button.pack(pady=20)

root.mainloop()
