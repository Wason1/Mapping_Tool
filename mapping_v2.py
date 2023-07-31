import pandas as pd
from fuzzywuzzy import fuzz, process
from tkinter import Tk, filedialog, Listbox, StringVar, END, messagebox, OptionMenu, Button, DISABLED, NORMAL
from tkinter.ttk import Progressbar
from tqdm import tqdm
import os

def load_spreadsheet(application, spreadsheet_number):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        df = pd.read_excel(filename)
        print(f'Spreadsheet {spreadsheet_number} loaded with shape {df.shape}')
        if spreadsheet_number == 1:
            application.spreadsheet1 = df
            application.update_dropdown(application.dropdown1, df.columns)
            application.load_button1.config(bg='blue')
            application.dropdown1.config(state=NORMAL)
            application.load_button2.config(state=NORMAL) # Enable "Load Spreadsheet 2" button right after Spreadsheet 1 is loaded
        elif spreadsheet_number == 2:
            application.spreadsheet2 = df
            application.update_dropdown(application.dropdown2, df.columns)
            application.load_button2.config(bg='blue')
            application.dropdown2.config(state=NORMAL)
        application.check_if_match_possible()

def match_data(s1, s2, progressbar):
    matches = []
    length = len(s1)
    progressbar["maximum"] = length
    for i in tqdm(range(length), desc="Matching..."):
        highest = process.extractOne(s1[i], s2, scorer=fuzz.token_sort_ratio)
        matches.append(highest)
        progressbar["value"] = i
        progressbar.update()
    return matches

class Application:
    def __init__(self, master):
        self.master = master
        self.spreadsheet1 = None
        self.spreadsheet2 = None
        self.column1 = None
        self.column2 = None
        self.matches = None
        self.listbox = Listbox(master)
        self.listbox.pack(fill='both', expand=True)

        self.variable1 = StringVar(master)
        self.variable1.set("Select column...")
        self.dropdown1 = OptionMenu(master, self.variable1, '')
        self.dropdown1.pack(side="left")
        self.dropdown1.config(state=DISABLED)

        self.variable2 = StringVar(master)
        self.variable2.set("Select column...")
        self.dropdown2 = OptionMenu(master, self.variable2, '')
        self.dropdown2.pack(side="right")
        self.dropdown2.config(state=DISABLED)

        self.load_button1 = Button(master, text="Load Spreadsheet 1", command=lambda: load_spreadsheet(self, 1))
        self.load_button1.pack(fill='x')
        self.load_button2 = Button(master, text="Load Spreadsheet 2", command=lambda: load_spreadsheet(self, 2), state=DISABLED)
        self.load_button2.pack(fill='x')
        self.match_button = Button(master, text="Match Data", command=self.match_data, state=DISABLED)
        self.match_button.pack(fill='x')
        self.save_button = Button(master, text="Save Matches", command=self.save_matches, state=DISABLED)
        self.save_button.pack(fill='x')
        
        self.progressbar = Progressbar(master, length=500)
        self.progressbar.pack(fill='x')

    def set_column(self, dropdown, value):
        if dropdown == self.dropdown1:
            self.column1 = value
        elif dropdown == self.dropdown2:
            self.column2 = value
        self.check_if_match_possible()

    def update_dropdown(self, dropdown, options):
        dropdown['menu'].delete(0, 'end')
        for option in options:
            dropdown['menu'].add_command(label=option, command=lambda value=option: self.set_column(dropdown, value))
            
    def check_if_match_possible(self):
        if self.spreadsheet1 is not None and self.spreadsheet2 is not None and self.column1 is not None and self.column2 is not None:
            self.match_button.config(state=NORMAL)

    def match_data(self):
        if self.spreadsheet1 is not None and self.spreadsheet2 is not None and self.column1 is not None and self.column2 is not None:
            self.matches = match_data(self.spreadsheet1[self.column1].values, self.spreadsheet2[self.column2].values, self.progressbar)
            self.listbox.delete(0, END)
            for match in self.matches:
                self.listbox.insert(END, match)
            self.match_button.config(bg='blue')
            self.save_button.config(state=NORMAL)
        else:
            messagebox.showerror("Error", "Please load both spreadsheets and select columns before matching data.")

    def save_matches(self):
        if self.matches:
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if filename:
                match_df = pd.DataFrame(self.matches)
                match_df.to_excel(filename, index=False)
                messagebox.showinfo("Success", "Matches saved successfully.")
                self.save_button.config(bg='blue')
        else:
            messagebox.showerror("Error", "No matches to save.")

root = Tk()
root.state('zoomed')  # To maximize the window
app = Application(root)
root.mainloop()
