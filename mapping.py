import pandas as pd
from fuzzywuzzy import fuzz, process
from tkinter import Tk, filedialog, Listbox, StringVar, END, messagebox, OptionMenu
from tkinter.ttk import Progressbar
from tqdm import tqdm
import os


def load_spreadsheet():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        df = pd.read_excel(filename)
        return df
    else:
        return None

def match_data(s1, s2, progressbar):
    matches = []
    length = len(s1)
    progressbar["maximum"] = length
    for i in tqdm(range(length), desc="Matching..."):
        highest = process.extractBests(s1[i], s2, scorer=fuzz.token_sort_ratio)
        matches.append(highest)
        progressbar["value"] = i
        progressbar.update()
    return matches

class Application:
    def __init__(self, master):
        self.master = master
        self.spreadsheet1 = None
        self.spreadsheet2 = None
        self.matches = None
        self.listbox = Listbox(master)
        self.listbox.pack()
        self.variable = StringVar(master)
        self.variable.set("Select an action...")
        self.dropdown = OptionMenu(master, self.variable, "Load Spreadsheet 1", "Load Spreadsheet 2", "Match Data", "Select Matches", "Save Matches", command=self.callback)
        self.dropdown.pack()
        self.progressbar = Progressbar(master, length=500)
        self.progressbar.pack()

    def callback(self, value):
        if value == "Load Spreadsheet 1":
            self.spreadsheet1 = load_spreadsheet()
            if self.spreadsheet1 is None:
                messagebox.showerror("Error", "Failed to load spreadsheet 1.")
        elif value == "Load Spreadsheet 2":
            self.spreadsheet2 = load_spreadsheet()
            if self.spreadsheet2 is None:
                messagebox.showerror("Error", "Failed to load spreadsheet 2.")
        elif value == "Match Data":
            if self.spreadsheet1 is not None and self.spreadsheet2 is not None:
                self.matches = match_data(self.spreadsheet1["NAME1"].values, self.spreadsheet2["NAME2"].values, self.progressbar)
                self.listbox.delete(0, END)
                for match in self.matches:
                    self.listbox.insert(END, match)
            else:
                messagebox.showerror("Error", "Please load both spreadsheets before matching data.")
        elif value == "Select Matches":
            self.selected = self.listbox.curselection()
            for index in self.selected:
                print(self.listbox.get(index))
        elif value == "Save Matches":
            if self.matches:
                filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
                if filename:
                    match_df = pd.DataFrame(self.matches)
                    match_df.to_excel(filename, index=False)
                    messagebox.showinfo("Success", "Matches saved successfully.")
            else:
                messagebox.showerror("Error", "No matches to save.")

root = Tk()
app = Application(root)
root.mainloop()
