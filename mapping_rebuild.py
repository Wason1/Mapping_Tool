import pandas as pd
import os
import ast
from fuzzywuzzy import fuzz, process
from tkinter import Tk, filedialog, StringVar, END, messagebox, OptionMenu, Button, DISABLED, NORMAL, Listbox, Checkbutton, Label, Frame
from tkinter.ttk import Progressbar
from tqdm import tqdm
import subprocess
import sys

# [ ... your imports remain unchanged ... ]

class Application:
    def __init__(self, master):
        self.master = master
        self.spreadsheet1 = None
        self.spreadsheet2 = None

        self.column1 = None
        self.column2 = None

        self.matches = []
        self.next_item_index = 0
        self.selections = {}

        # Top frame for the buttons
        self.top_frame = Frame(master)
        self.top_frame.pack(fill='x', pady=10)

        self.load_button1 = Button(self.top_frame, text="Load Spreadsheet 1", command=lambda: self.load_spreadsheet(1), bg='green')
        # self.load_button1.grid(row=0, column=0, sticky='ew', padx=5)

        self.load_button2 = Button(self.top_frame, text="Load Spreadsheet 2", command=lambda: self.load_spreadsheet(2), state=DISABLED)
        # self.load_button2.grid(row=1, column=0, sticky='ew', padx=5)

        self.variable1 = StringVar(master)
        self.variable1.set("Select matching column from 1...")
        self.dropdown1 = OptionMenu(self.top_frame, self.variable1, '')
        self.dropdown1.config(state=DISABLED)
        # self.dropdown1.grid(row=0, column=1, sticky='ew', padx=5)

        self.variable2 = StringVar(master)
        self.variable2.set("Select matching column from 2...")
        self.dropdown2 = OptionMenu(self.top_frame, self.variable2, '')
        self.dropdown2.config(state=DISABLED)
        # self.dropdown2.grid(row=1, column=1, sticky='ew', padx=5)

        self.match_button = Button(self.top_frame, text="Initiate Mapping Process", command=self.start_matching, state=DISABLED)
        # self.match_button.grid(row=2, column=0, sticky='ew', padx=5)

        self.next_button = Button(self.top_frame, text="Next Item", command=self.next_item, state=DISABLED)
        # self.next_button.grid(row=2, column=1, sticky='ew', padx=5)

        # Top frame grid
        self.load_button1.grid(row=0, column=0, sticky='ew', padx=5, columnspan=2)
        self.load_button2.grid(row=1, column=0, sticky='ew', padx=5, columnspan=2)
        self.dropdown1.grid(row=0, column=2, sticky='ew', padx=5)
        self.dropdown2.grid(row=1, column=2, sticky='ew', padx=5)
        self.match_button.grid(row=2, column=0, sticky='ew', padx=5, columnspan=2)
        self.next_button.grid(row=2, column=2, sticky='ew', padx=5)

        self.top_frame.grid_columnconfigure(0, weight=1)
        self.top_frame.grid_columnconfigure(1, weight=1)
        self.top_frame.grid_columnconfigure(2, weight=1)

        # Middle frame for future content with a border
        self.middle_frame = Frame(master, bd=1, relief='solid')
        self.middle_frame.pack(fill='both', expand=True, pady=10)

        # Bottom Frame for reset, close, progress bar
        self.bottom_frame = Frame(master)
        self.bottom_frame.pack(fill='x', side='bottom', pady=10)

        # Progress Bar
        self.progressbar = Progressbar(self.bottom_frame, length=500)
        self.progressbar.pack(fill='x')

        self.progress_label = Label(self.bottom_frame, text="")
        self.progress_label.pack(fill='x')

        self.refresh_button = Button(self.bottom_frame, text="Reset", command=self.refresh)
        self.refresh_button.pack(fill='x')

        self.save_button = Button(self.bottom_frame, text="Save Matches", command=lambda: self.save_selections(self.spreadsheet1, self.spreadsheet2), state=DISABLED)
        self.save_button.pack(fill='x')

        self.close_button = Button(self.bottom_frame, text="Close", command=self.close_app)
        self.close_button.pack(fill='x')

        self.match_frame = Frame(self.middle_frame)  # Container for match widgets
        self.match_frame.pack(fill='both', expand=True, padx=10, pady=10)




    def load_spreadsheet(self, spreadsheet_number):
        filepath = filedialog.askopenfilename(title=f"Open Spreadsheet {spreadsheet_number}", filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")))
        if filepath:
            try:
                if filepath.endswith('.csv'):
                    df = pd.read_csv(filepath)
                else:
                    df = pd.read_excel(filepath)

                if spreadsheet_number == 1:
                    self.spreadsheet1 = df
                    self.load_button1.config(bg="blue", text=filepath)

                    # Activate the 'load spreadsheet 2' button
                    self.load_button2.config(state=NORMAL)  # This line makes the dropdown 2 button active

                    # Set dropdown column options for spreadsheet 1
                    self.update_dropdown(self.dropdown1, df.columns)

                elif spreadsheet_number == 2:
                    self.spreadsheet2 = df
                    self.load_button2.config(bg="blue", text=filepath)
                    # Set dropdown column options for spreadsheet 2
                    self.update_dropdown(self.dropdown2, df.columns)
                    #Activate dropdown button 2
                    self.dropdown2.config(state=NORMAL)
                    #activate dropdown button 1
                    self.dropdown1.config(state=NORMAL)

            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while loading the file:\n{e}")
    
    
    def update_dropdown(self, dropdown, options):
        dropdown['menu'].delete(0, 'end')
        for option in options:
            dropdown['menu'].add_command(label=option, command=lambda value=option: self.set_column(dropdown, value))


    def set_column(self, dropdown, value):
        # Existing logic
        if dropdown == self.dropdown1:
            self.column1 = value
            #column_button.config(text=f"Spreadsheet 1: {value}", bg="blue")
            self.dropdown1.config(text=f"Spreadsheet 1: {value}", bg="blue")
            self.dropdown2.config(state=NORMAL, bg='green')
        elif dropdown == self.dropdown2:
            self.column2 = value
            #column_button.config(text=f"Spreadsheet 2: {value}", bg="blue")
            self.dropdown2.config(text=f"Spreadsheet 2: {value}", bg="blue")
            self.match_button.config(state=NORMAL, bg='green')  # Enable "Start Matchin" button right after Spreadsheet 2 is loaded
        # Update the selected column display
        if dropdown == self.dropdown1:
            self.column1 = value
            self.variable1.set(value)
        elif dropdown == self.dropdown2:
            self.column2 = value
            self.variable2.set(value)
            
    def start_matching(self):
        self.progressbar["maximum"] = len(self.spreadsheet1)
        self.progress_label.config(text="0%")
        self.match_button.config(bg='blue', state=DISABLED)
        self.next_button.config(state=NORMAL)
        self.save_button.config(state=NORMAL)

    def next_item(self):
        row_df = self.spreadsheet1.iloc[[self.next_item_index]]
        self.next_item_index += 1
        self.next_button.config(text="Next Item")

    def save_selections(self, df1, df2):
        # Extract the selections from the checkbox items
        selected_matches = []
        selected_index_matches = []
        key_index_number = int(0)
        for key, value in self.selections.items():
            s1_match = key
            for var in value:
                s2_match = var.get()
                if s2_match:
                    s2_match_tuple = ast.literal_eval(s2_match)
                    s2_match_index = int(s2_match_tuple[0])
                    selected_matches.append((s1_match, s2_match))
                    selected_index_matches.append([key_index_number, s2_match_index])
            key_index_number += 1

        print("FINAL SELECTED MATCHES: ",selected_matches)
        print("FINAL SELECTED INDEXES: ",selected_index_matches)


        # Initialize empty DataFrame
        df_joined = pd.DataFrame()

        # Loop through the selected index matches
        for i, j in selected_index_matches:
            # Extract the corresponding rows
            row_df1 = df1.iloc[[i]]
            row_df1 = row_df1.reset_index(drop=True)
            print(row_df1)
            row_df2 = df2.iloc[[j]]
            row_df2 = row_df2.reset_index(drop=True)
            print(row_df2)
            # Concatenate the rows
            row_joined = pd.concat([row_df1, row_df2], axis=1)
            print(row_joined)
            # Append the result to df_joined
            #df_joined = df_joined.append(row_joined)
            df_joined = pd.concat([df_joined, row_joined], ignore_index=True)
            print(df_joined)


        # Assuming df1 and df2 have same column names
        cols = pd.Series(df_joined.columns)
        for dup in cols[cols.duplicated()].unique(): 
            cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
        df_joined.columns = cols

        # Save the selected matches to an Excel file
        if selected_matches:
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if filename:
                match_df = pd.DataFrame(selected_matches, columns=['Name1', 'Name2'])
                df_joined.to_excel(filename, index=False)
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
        self.selected_column1_label.config(text="")
        self.selected_column2_label.config(text="")

        # Reset buttons
        self.load_button1.config(text="Load Spreadsheet 1", bg='green', state=NORMAL)
        self.load_button2.config(text="Load Spreadsheet 2", state=DISABLED, bg='SystemButtonFace')
        self.match_button.config(text="Match Data", state=DISABLED, bg='SystemButtonFace')
        self.next_button.config(text="First Item to Map", state=DISABLED, bg='SystemButtonFace')
        self.save_button.config(text="Save Matches", state=DISABLED, bg='SystemButtonFace')
        
        # Reset progress bar
        self.progressbar["value"] = 0
        self.progress_label.config(text="0%")  # Add this line

        
        # Clear the match_frame
        for widget in self.match_frame.winfo_children():
            widget.destroy()

 



root = Tk()
root.state('zoomed')  # To maximize the window
app = Application(root)
root.mainloop()