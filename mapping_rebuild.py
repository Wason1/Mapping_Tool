import pandas as pd
import os
import ast
from fuzzywuzzy import fuzz, process
from tkinter import Tk, filedialog, StringVar, END, messagebox, OptionMenu, Button, DISABLED, NORMAL, Listbox, Checkbutton, Label, Frame
from tkinter.ttk import Progressbar
from tqdm import tqdm
import subprocess
import sys



class Application:
    def __init__(self, master):
        self.master = master
        self.spreadsheet1 = None
        self.spreadsheet2 = None

        self.column1 = None
        self.column2 = None

        self.load_button1 = Button(master, text="Load Spreadsheet 1", command=lambda: self.load_spreadsheet(1), bg='green')
        self.load_button1.pack(fill='x')

        self.load_button2 = Button(master, text="Load Spreadsheet 2", command=lambda: self.load_spreadsheet(2), state=DISABLED)
        self.load_button2.pack(fill='x')

        self.variable1 = StringVar(master)
        self.variable1.set("Select matching column from 1...")
        self.dropdown1 = OptionMenu(master, self.variable1, '')
        self.dropdown1.pack()
        self.dropdown1.config(state=DISABLED)

        self.variable2 = StringVar(master)
        self.variable2.set("Select matching column from 2...")
        self.dropdown2 = OptionMenu(master, self.variable2, '')
        self.dropdown2.pack()
        self.dropdown2.config(state=DISABLED)

        self.selected_column1_label = Label(master, text="")
        self.selected_column1_label.pack()
        self.selected_column2_label = Label(master, text="")
        self.selected_column2_label.pack()


        self.match_button = Button(master, text="Match Data", command=self.match_data, state=DISABLED)
        self.match_button.pack(fill='x')

        self.next_button = Button(master, text="Next Item", command=self.next_item, state=DISABLED)
        self.next_button.pack(fill='x')

        self.save_button = Button(master, text="Save Matches", command=lambda: self.save_selections(self.spreadsheet1, self.spreadsheet2), state=DISABLED)
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
                    self.load_button1.config(bg="blue")

                    # Activate the 'load spreadsheet 2' button
                    self.load_button2.config(state=NORMAL)  # This line makes the button active

                elif spreadsheet_number == 2:
                    self.spreadsheet2 = df
                    self.load_button2.config(bg="blue")

                    #Activate dropdown 2
                    self.update_dropdown(self.dropdown2, df.columns)
                    self.dropdown2.config(state=NORMAL)

                    #activate dropdown 1
                    self.update_dropdown(self.dropdown1, df.columns)
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
        self.match_button.config(state=NORMAL, bg='green')  # Enable "Match Data" button right after Spreadsheet 2 is loaded
    # Update the selected column display
    if dropdown == self.dropdown1:
        self.column1 = value
        self.selected_column1_label.config(text=value)
    elif dropdown == self.dropdown2:
        self.column2 = value
        self.selected_column2_label.config(text=value)


    def set_column(self, dropdown, value):

            
    def prepare_match_data(self, df1, df2, column1, column2, progressbar, threshold=30):
        self.df1 = df1  # store the dataframes for future reference
        self.df2 = df2
        self.column1 = column1
        self.column2 = column2
        self.threshold = threshold
        self.progressbar = progressbar
        
        self.progressbar["maximum"] = len(self.df1)
        self.matches = []  # Reset matches list
        self.next_item_index = 0  # Reset index

        # Process only the first item initially
        self.next_item()

    def match_data(self):
        if self.spreadsheet1 is not None and self.spreadsheet2 is not None and self.column1 is not None and self.column2 is not None:
            self.prepare_match_data(self.spreadsheet1, self.spreadsheet2, self.column1, self.column2, self.progressbar)
            self.match_button.config(bg='blue', state=DISABLED)
            self.save_button.config(state=NORMAL)  # Enable "Save Matches" button right after matching is complete
        else:
            messagebox.showerror("Error", "Please load both spreadsheets and select columns before matching data.")
    
    def generate_match(self, index):
        """Generate match for the item at the provided index."""
        s1_item = self.df1[self.column1].iloc[index]
        s2_with_indices = [(i, elem) for i, elem in enumerate(self.df2[self.column2])]
        matches = process.extract(s1_item, s2_with_indices, scorer=fuzz.token_sort_ratio, limit=20)

        good_matches = [pd.DataFrame({'Index': [match[0][0]], 
                                      'Match': [match[0][1]], 
                                      'Score': [match[1]]}) for match in matches if match[1] >= self.threshold]

        self.matches.append((self.df1.iloc[index], pd.concat(good_matches, ignore_index=True)))
        
        # Update progress bar
        self.progressbar["value"] = index + 1
        self.progressbar.update()

    def next_item(self):
        # Check if we need to generate the next match
        if self.next_item_index >= len(self.matches) and self.next_item_index < len(self.df1):
            self.generate_match(self.next_item_index)

        if self.matches and self.next_item_index < len(self.matches):
            # Clear the match_frame
            for widget in self.match_frame.winfo_children():
                widget.destroy()

            row, matches_df = self.matches[self.next_item_index]
            self.match_frame.pack_forget()

            Label(self.match_frame, text=f"Matches for row:\n{row.to_string()}").pack()

            # Create a list to store the StringVar objects for this item
            self.selections[row[self.column1]] = []

            for _, match_row in matches_df.iterrows():
                match_text = f"Index: {match_row['Index']}, Match: {match_row['Match']}, Score: {match_row['Score']}"
                var = StringVar()
                Checkbutton(self.match_frame, text=match_text, variable=var, onvalue=match_text, offvalue="").pack()
                self.selections[row[self.column1]].append(var)

            self.next_item_index += 1  # Prepare for the next click of the "Next Item" button
            self.next_button.config(state=NORMAL if self.next_item_index < len(self.df1) else DISABLED)

            self.match_frame.pack(fill='both', expand=True)

        else:
            messagebox.showinfo("Info", "No more items to match.")

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

        # Reset buttons
        self.load_button1.config(text="Load Spreadsheet 1", bg='green', state=NORMAL)
        self.load_button2.config(text="Load Spreadsheet 2", state=DISABLED, bg='SystemButtonFace')
        self.match_button.config(text="Match Data", state=DISABLED, bg='SystemButtonFace')
        self.next_button.config(text="Next Item", state=DISABLED, bg='SystemButtonFace')
        self.save_button.config(text="Save Matches", state=DISABLED, bg='SystemButtonFace')
        
        # Reset progress bar
        self.progressbar["value"] = 0
        
        # Clear the match_frame
        for widget in self.match_frame.winfo_children():
            widget.destroy()

 



root = Tk()
root.state('zoomed')  # To maximize the window
app = Application(root)
root.mainloop()