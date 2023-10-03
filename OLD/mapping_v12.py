import pandas as pd
import os
from fuzzywuzzy import fuzz, process
from tkinter import Tk, filedialog, StringVar, END, messagebox, OptionMenu, Button, DISABLED, NORMAL, Listbox, Checkbutton, Label, Frame
from tkinter.ttk import Progressbar
from tqdm import tqdm
import subprocess
import sys

def load_spreadsheet(application, spreadsheet_number):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        if spreadsheet_number == 1:
            df_1 = pd.read_excel(filename)
            print(f'Spreadsheet {spreadsheet_number} loaded with shape {df_1.shape}')
            application.spreadsheet1 = df_1
            application.update_dropdown(application.dropdown1, df_1.columns)
            application.load_button1.config(bg='blue')
            application.dropdown1.config(state=NORMAL)
            application.load_button2.config(state=NORMAL, bg='green') # Enable "Load Spreadsheet 2" button right after Spreadsheet 1 is loaded
        elif spreadsheet_number == 2:
            df_2 = pd.read_excel(filename)
            print(f'Spreadsheet {spreadsheet_number} loaded with shape {df_2.shape}')
            application.spreadsheet2 = df_2
            application.update_dropdown(application.dropdown2, df_2.columns)
            application.load_button2.config(bg='blue')
            application.dropdown2.config(state=NORMAL)
            application.match_button.config(state=NORMAL, bg='green')  # Enable "Match Data" button right after Spreadsheet 2 is loaded

def match_data(application, s1, s2, progressbar, threshold=70):
    length = len(s1)
    progressbar["maximum"] = length
    application.matches = []  # Create a list to store all matches
    for i in tqdm(range(length), desc="Matching..."):
        matches = process.extract(s1[i], s2, scorer=fuzz.token_sort_ratio, limit=None)
        good_matches = [match for match in matches if match[1] >= threshold]
        application.matches.append((s1[i], good_matches))
        progressbar["value"] = i
        progressbar.update()
    application.next_item_index = 0  # Initialize index for "Next Item" button
    application.next_item()  # Display the first item

def save_selections(application):
    # Extract the selections from the checkbox items
    selected_matches = []
    for key, value in application.selections.items():
        s1_match = key
        for var in value:
            s2_match = var.get()
            if s2_match:
                selected_matches.append((s1_match, s2_match))

    # Save the selected matches to an Excel file
    if selected_matches:
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            match_df = pd.DataFrame(selected_matches, columns=['Name1', 'Name2'])
            match_df.to_excel(filename, index=False)
            messagebox.showinfo("Success", "Matches saved successfully.")
            
            # Use the subprocess module to open the file with the default application
            if sys.platform.startswith('darwin'):  # macOS
                subprocess.call(('open', filename))
            elif sys.platform.startswith('linux'):  # linux
                subprocess.call(('xdg-open', filename))
            else:  # windows
                os.startfile(filename)
    else:
        messagebox.showerror("Error", "No matches to save.")


  
class Application:
    def __init__(self, master):
        self.master = master
        self.spreadsheet1 = None
        self.spreadsheet2 = None

        self.column1 = None
        self.column2 = None

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

        self.load_button1 = Button(master, text="Load Spreadsheet 1", command=lambda: load_spreadsheet(self, 1), bg='green')
        self.load_button1.pack(fill='x')

        self.load_button2 = Button(master, text="Load Spreadsheet 2", command=lambda: load_spreadsheet(self, 2), state=DISABLED)
        self.load_button2.pack(fill='x')

        #self.column_button1 = Button(master, text="Select column from Spreadsheet 1", state=DISABLED)
        #self.column_button1.pack(fill='x')

        #self.column_button2 = Button(master, text="Select column from Spreadsheet 2", state=DISABLED)
        #self.column_button2.pack(fill='x')

        self.match_button = Button(master, text="Match Data", command=self.match_data, state=DISABLED)
        self.match_button.pack(fill='x')

        self.next_button = Button(master, text="Next Item", command=self.next_item, state=DISABLED)
        self.next_button.pack(fill='x')

        #self.save_button = Button(master, text="Save Matches", command=self.save_matches, state=DISABLED)
        self.save_button = Button(master, text="Save Matches", command=lambda: save_selections(self), state=DISABLED)
        self.save_button.pack(fill='x')

        self.close_button = Button(master, text="Close", command=self.close_app)
        self.close_button.pack(fill='x')

        self.progressbar = Progressbar(master, length=500)
        self.progressbar.pack(fill='x')

        self.refresh_button = Button(master, text="Reset", command=self.refresh)
        self.refresh_button.pack(fill='x')


        self.match_frame = Frame(master)  # Container for match widgets
        self.match_frame.pack(fill='both', expand=True)


        self.matches = []
        self.next_item_index = 0
        self.selections = {}
    
    def close_app(self):
        self.master.destroy()

    def refresh(self):
        # Reset variables
        self.spreadsheet1 = None
        self.spreadsheet2 = None
        self.column1 = None
        self.column2 = None
        self.matches = []
        self.next_item_index = 0
        self.selections = {}

        # Reset dropdowns
        self.variable1.set("Select column...")
        self.variable2.set("Select column...")
        self.dropdown1.config(state=DISABLED, bg='SystemButtonFace')
        self.dropdown2.config(state=DISABLED, bg='SystemButtonFace')
        self.dropdown1['menu'].delete(0, 'end')
        self.dropdown2['menu'].delete(0, 'end')

        # Reset buttons
        self.load_button1.config(text="Load Spreadsheet 1", bg='green', state=NORMAL)
        self.load_button2.config(text="Load Spreadsheet 2", state=DISABLED, bg='SystemButtonFace')
        #self.column_button1.config(text="Select column from Spreadsheet 1", state=DISABLED, bg='SystemButtonFace')
        #self.column_button2.config(text="Select column from Spreadsheet 2", state=DISABLED, bg='SystemButtonFace')
        self.match_button.config(text="Match Data", state=DISABLED, bg='SystemButtonFace')
        self.next_button.config(text="Next Item", state=DISABLED, bg='SystemButtonFace')
        self.save_button.config(text="Save Matches", state=DISABLED, bg='SystemButtonFace')
        
        # Reset progress bar
        self.progressbar["value"] = 0
        
        # Clear the match_frame
        for widget in self.match_frame.winfo_children():
            widget.destroy()


    def set_column(self, dropdown, value):
        if dropdown == self.dropdown1:
            self.column1 = value
            self.load_button1.config(text=f"Spreadsheet 1: {value}", bg="blue")
        elif dropdown == self.dropdown2:
            self.column2 = value
            self.load_button2.config(text=f"Spreadsheet 2: {value}", bg="blue")
        

    def update_dropdown(self, dropdown, options):
        dropdown['menu'].delete(0, 'end')
        for option in options:
            dropdown['menu'].add_command(label=option, command=lambda value=option: self.set_column(dropdown, value))

    def set_column(self, dropdown, value):
        if dropdown == self.dropdown1:
            self.column1 = value
            #column_button.config(text=f"Spreadsheet 1: {value}", bg="blue")
            self.dropdown1.config(text=f"Spreadsheet 1: {value}", bg="blue")
        elif dropdown == self.dropdown2:
            self.column2 = value
            #column_button.config(text=f"Spreadsheet 2: {value}", bg="blue")
            self.dropdown2.config(text=f"Spreadsheet 2: {value}", bg="blue")
        

    def match_data(self):
        if self.spreadsheet1 is not None and self.spreadsheet2 is not None and self.column1 is not None and self.column2 is not None:
            match_data(self, self.spreadsheet1[self.column1].values, self.spreadsheet2[self.column2].values, self.progressbar)
            self.match_button.config(bg='blue')
            self.save_button.config(state=NORMAL)  # Enable "Save Matches" button right after matching is complete
        else:
            messagebox.showerror("Error", "Please load both spreadsheets and select columns before matching data.")

    def add_match(self, match):  # Add the match to the listbox instead of a dictionary
        self.listbox.insert(END, str(match))

    def save_matches(self):
        matches = [self.listbox.get(idx) for idx in range(self.listbox.size())]  # Get all the matches from the listbox
        if matches:
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if filename:
                match_df = pd.DataFrame(matches, columns=["Match"])
                match_df.to_excel(filename, index=False)
                messagebox.showinfo("Success", "Matches saved successfully.")
        else:
            messagebox.showerror("Error", "No matches to save.")
    
    def next_item(self):
        if self.matches and self.next_item_index < len(self.matches):
            # Clear the match_frame
            for widget in self.match_frame.winfo_children():
                widget.destroy()
                
            item, matches = self.matches[self.next_item_index]
            self.match_frame.pack_forget()
            
            Label(self.match_frame, text=f"Matches for '{item}':").pack()

            # Create a list to store the StringVar objects for this item
            self.selections[item] = []
            
            for match in matches:
                var = StringVar()
                Checkbutton(self.match_frame, text=str(match), variable=var, onvalue=str(match), offvalue="").pack()
                self.selections[item].append(var)

            self.next_item_index += 1  # Prepare for the next click of the "Next Item" button
            self.next_button.config(state=NORMAL if self.next_item_index < len(self.matches) else DISABLED)

            self.progressbar["value"] = self.next_item_index
            self.progressbar.update()
            
            self.match_frame.pack(fill='both', expand=True)

        else:
            messagebox.showinfo("Info", "No more items to match.")


root = Tk()
root.state('zoomed')  # To maximize the window
app = Application(root)
root.mainloop()