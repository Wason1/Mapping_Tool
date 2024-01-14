import pandas as pd
import os
import ast
from fuzzywuzzy import fuzz, process
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, StringVar, END, messagebox, OptionMenu, Button, DISABLED, NORMAL, Checkbutton, Label, Frame, Canvas, Scrollbar
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
        self.matches = []
        self.next_item_index = 0
        self.selections = {}
        self.max_index = int(1)
        # Current Item you're mapping
        self.current_item_var = StringVar()

        # Top frame for the buttons
        #region
        self.top_frame = Frame(master)
        self.top_frame.pack(fill='x', pady=10)
        self.load_button1 = Button(self.top_frame, text="Load Spreadsheet 1", command=lambda: self.load_spreadsheet(1), bg='green')
        self.load_button2 = Button(self.top_frame, text="Load Spreadsheet 2", command=lambda: self.load_spreadsheet(2), state=DISABLED)
        self.variable1 = StringVar(master)
        self.variable1.set("Select matching column from 1...")
        self.dropdown1 = OptionMenu(self.top_frame, self.variable1, '')
        self.dropdown1.config(state=DISABLED)
        self.variable2 = StringVar(master)
        self.variable2.set("Select matching column from 2...")
        self.dropdown2 = OptionMenu(self.top_frame, self.variable2, '')
        self.dropdown2.config(state=DISABLED)
        self.match_button = Button(self.top_frame, text="Initiate Mapping Process", command=self.start_matching, state=DISABLED)
        self.next_button = Button(self.top_frame, text="Map First Item", command=self.next_item, state=DISABLED)
        
        # Top frame grid
        self.load_button1.grid(row=0, column=0, sticky='ew', padx=5, columnspan=2)
        self.load_button2.grid(row=1, column=0, sticky='ew', padx=5, columnspan=2)
        self.dropdown1.grid(row=0, column=2, sticky='ew', padx=5)
        self.dropdown2.grid(row=1, column=2, sticky='ew', padx=5)
        self.match_button.grid(row=2, column=0, sticky='ew', padx=5, columnspan=2)
        self.next_button.grid(row=2, column=2, sticky='ew', padx=5)
        # self.current_item_label.grid(row=3, column=1, sticky='w', padx=5)
        self.top_frame.grid_columnconfigure(0, weight=1)
        self.top_frame.grid_columnconfigure(1, weight=1)
        self.top_frame.grid_columnconfigure(2, weight=1)
        #endregion



        #Curent Item Frame
        #region
        self.current_item_frame = Frame(master, bd=1, relief='solid')
        self.current_item_frame.grid_propagate(False)
        self.current_item_frame.pack(fill='both', expand=True, pady=10)
        #Left Frame
        self.current_item_left_frame = Frame(self.current_item_frame, bd=1, relief='solid')
        self.current_item_left_frame.pack(fill='both', expand=True, side='left', padx=(5, 10), pady=10)
        # self.current_item_left_frame.pack(fill='both', expand=True, side='left', padx=(5, 2.5), pady=10)
        self.current_item_left_canvas = Canvas(self.current_item_left_frame)
        self.current_item_left_scrollbar = Scrollbar(self.current_item_left_frame, orient="vertical", command=self.current_item_left_canvas.yview)
        self.current_item_left_scrollbar_horizontal = Scrollbar(self.current_item_left_frame, orient="horizontal", command=self.current_item_left_canvas.xview)
        # self.current_item_left_canvas.config(scrollregion=self.current_item_left_canvas.bbox("all"))
        self.current_item_left_canvas.configure(yscrollcommand=self.current_item_left_scrollbar.set, xscrollcommand=self.current_item_left_scrollbar_horizontal.set)
        self.current_item_left_canvas.grid(row=0, column=0, sticky="nsew")
        self.current_item_left_scrollbar.grid(row=0, column=1, sticky="ns")
        self.current_item_left_scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.current_item_left_frame.grid_rowconfigure(0, weight=1)
        self.current_item_left_frame.grid_columnconfigure(0, weight=1)
        # Create a frame to hold the Treeview, and place it within current_item_left_inner_frame
        self.current_item_left_inner_frame = Frame(self.current_item_left_canvas)
        self.df_frame = Frame(self.current_item_left_inner_frame)
        self.df_frame.pack(fill='both', expand=True)
        # Create a Treeview widget
        self.tree = ttk.Treeview(self.df_frame)
        self.tree.pack(fill='both', expand=True)
        self.current_item_left_canvas.create_window((0, 0), window=self.current_item_left_inner_frame, anchor='nw')
        self.current_item_left_inner_frame.bind('<Configure>', lambda e: self.current_item_left_canvas.configure(scrollregion=self.current_item_left_canvas.bbox("all")))        
        #Right Frame
        self.current_item_right_frame = Frame(self.current_item_frame, bd=1, relief='solid')
        self.current_item_right_frame.pack(fill='both', expand=True, side='right', padx=(10, 5), pady=10)
        self.current_item_right_canvas = Canvas(self.current_item_right_frame)
        self.current_item_right_scrollbar = Scrollbar(self.current_item_right_frame, orient="vertical", command=self.current_item_right_canvas.yview)
        self.current_item_right_scrollbar_horizontal = Scrollbar(self.current_item_right_frame, orient="horizontal", command=self.current_item_right_canvas.xview)
        self.current_item_right_canvas.configure(yscrollcommand=self.current_item_right_scrollbar.set, xscrollcommand=self.current_item_right_scrollbar_horizontal.set)        
        self.current_item_right_canvas.grid(row=0, column=0, sticky="nsew")
        self.current_item_right_scrollbar.grid(row=0, column=1, sticky="ns")
        self.current_item_right_scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.current_item_right_frame.grid_rowconfigure(0, weight=1)
        self.current_item_right_frame.grid_columnconfigure(0, weight=1)
        self.current_item_right_inner_frame = Frame(self.current_item_right_canvas)
        self.current_item_right_label = Label(self.current_item_right_inner_frame, textvariable=self.current_item_var)
        self.current_item_right_label.pack() 
        self.current_item_right_canvas.create_window((0, 0), window=self.current_item_right_inner_frame, anchor='nw')
        self.current_item_right_inner_frame.bind('<Configure>', lambda e: self.current_item_right_canvas.configure(scrollregion=self.current_item_right_canvas.bbox("all")))

        # self.current_item_right_canvas.bind('<Configure>', lambda e: self.current_item_right_canvas.configure(scrollregion=self.current_item_right_canvas.bbox("all")))

        #endregion

        

        # Middle frame
        #region
        # Middle frame with a border
        self.middle_frame = Frame(master, bd=1, relief='solid')
        self.middle_frame.pack(fill='both', expand=True, pady=10)

        # Middle frame's left and right frames
        self.middle_left_frame = Frame(self.middle_frame, bd=1, relief='solid')
        self.middle_left_frame.pack(fill='both', expand=True, side='left', padx=(5, 0), pady=10)
        # self.middle_left_frame.pack(fill='both', expand=True, side='left', padx=(5, 2.5), pady=10)
        self.middle_right_frame = Frame(self.middle_frame, bd=1, relief='solid')
        self.middle_right_frame.pack(fill='both', expand=True, side='right', padx=(0, 5), pady=10)
        # self.middle_right_frame.pack(fill='both', expand=True, side='right', padx=(2.5, 5), pady=10)

        # Canvas and Scrollbars for middle_right_frame
        self.middle_right_canvas = Canvas(self.middle_right_frame)
        self.middle_right_scrollbar = Scrollbar(self.middle_right_frame, orient="vertical", command=self.middle_right_canvas.yview)
        self.middle_right_scrollbar_horizontal = Scrollbar(self.middle_right_frame, orient="horizontal", command=self.middle_right_canvas.xview)
        self.middle_right_canvas.configure(yscrollcommand=self.middle_right_scrollbar.set, xscrollcommand=self.middle_right_scrollbar_horizontal.set)

        # Grid layout for right canvas and scrollbars
        self.middle_right_canvas.grid(row=0, column=0, sticky="nsew")
        self.middle_right_scrollbar.grid(row=0, column=1, sticky="ns")
        self.middle_right_scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.middle_right_frame.grid_rowconfigure(0, weight=1)
        self.middle_right_frame.grid_columnconfigure(0, weight=1)

        # Inner Frame for middle_right_canvas
        self.middle_right_inner_frame = Frame(self.middle_right_canvas)
        self.middle_right_canvas.create_window((0, 0), window=self.middle_right_inner_frame, anchor='nw')

        # Canvas and Scrollbars for middle_left_frame
        self.middle_left_canvas = Canvas(self.middle_left_frame)
        self.middle_left_scrollbar = Scrollbar(self.middle_left_frame, orient="vertical", command=self.middle_left_canvas.yview)
        self.middle_left_scrollbar_horizontal = Scrollbar(self.middle_left_frame, orient="horizontal", command=self.middle_left_canvas.xview)
        self.middle_left_canvas.configure(yscrollcommand=self.middle_left_scrollbar.set, xscrollcommand=self.middle_left_scrollbar_horizontal.set)

        # Grid layout for left canvas and scrollbars
        self.middle_left_canvas.grid(row=0, column=0, sticky="nsew")
        self.middle_left_scrollbar.grid(row=0, column=1, sticky="ns")
        self.middle_left_scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.middle_left_frame.grid_rowconfigure(0, weight=1)
        self.middle_left_frame.grid_columnconfigure(0, weight=1)

        # Inner Frame for middle_left_canvas
        self.middle_left_inner_frame = Frame(self.middle_left_canvas)
        self.middle_left_canvas.create_window((0, 0), window=self.middle_left_inner_frame, anchor='nw')

        # Configure the canvas scroll region whenever the inner frame size changes.
        self.middle_left_inner_frame.bind('<Configure>', lambda e: self.middle_left_canvas.configure(scrollregion=self.middle_left_canvas.bbox("all")))
        self.middle_right_inner_frame.bind('<Configure>', lambda e: self.middle_right_canvas.configure(scrollregion=self.middle_right_canvas.bbox("all")))
        

        #endregion

        # Bottom Frame for reset, close, progress bar
        #region
        self.bottom_frame = Frame(master)
        self.bottom_frame.pack(fill='x', side='bottom', pady=10)
        # Progress Bar
        self.progressbar = Progressbar(self.bottom_frame, length=500)
        self.progressbar.pack(fill='x')
        self.progress_label = Label(self.bottom_frame, text="")
        self.progress_label.pack(fill='x')
        self.save_button = Button(self.bottom_frame, text="Save Matches", command=self.save_selections, state=DISABLED)
        self.save_button.pack(fill='x')
        self.close_button = Button(self.bottom_frame, text="Close", command=self.close_app)
        self.close_button.pack(fill='x')
        self.match_frame = Frame(self.middle_frame)  # Container for match widgets
        self.match_frame.pack(fill='both', expand=True, padx=10, pady=10)
        #endregion

    def load_spreadsheet(self, spreadsheet_number):
        filepath = filedialog.askopenfilename(title=f"Open Spreadsheet {spreadsheet_number}", filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")))
        if filepath:
            try:
                if filepath.endswith('.csv'):
                    df = pd.read_csv(filepath)
                else:
                    df = pd.read_excel(filepath)
                if spreadsheet_number == 1:
                    df.columns = [col + '_1' for col in df.columns]
                    self.spreadsheet1 = df
                    self.load_button1.config(bg="white", text=filepath, state=DISABLED)
                    # Activate the 'load spreadsheet 2' button
                    self.load_button2.config(state=NORMAL, bg="green")  # This line makes the dropdown 2 button active
                    # Set dropdown column options for spreadsheet 1
                    self.update_dropdown(self.dropdown1, df.columns)
                elif spreadsheet_number == 2:
                    df.columns = [col + '_2' for col in df.columns]
                    self.spreadsheet2 = df
                    self.load_button2.config(bg="white", text=filepath, state=DISABLED)
                    # Set dropdown column options for spreadsheet 2
                    self.update_dropdown(self.dropdown2, df.columns)
                    #activate dropdown button 1
                    self.dropdown1.config(state=NORMAL, bg="green")
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
            #column_button.config(text=f"Spreadsheet 1: {value}", bg="white")
            self.dropdown1.config(text=f"Spreadsheet 1: {value}", bg="white", state=DISABLED)
            self.dropdown2.config(state=NORMAL, bg='green')
        elif dropdown == self.dropdown2:
            self.column2 = value
            #column_button.config(text=f"Spreadsheet 2: {value}", bg="white")
            self.dropdown2.config(text=f"Spreadsheet 2: {value}", bg="white", state=DISABLED)
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
        self.match_button.config(bg='white', state=DISABLED)
        self.next_button.config(state=NORMAL, bg="green")
        self.save_button.config(state=NORMAL)
        messagebox.showinfo("Mapping", f"You are going to map {len(self.spreadsheet1)} items")
        self.max_index = len(self.spreadsheet1) - 1
        # Generate a series for the column to map against
        self.matching_data_series = self.spreadsheet2[self.column2]
        self.df_final = pd.DataFrame()

    def fuzzy_logic_dataframe(self, input_string, series):
        # Calculate similarity scores
        sort_scores = series.apply(lambda x: fuzz.token_sort_ratio(input_string, x))
        set_scores = series.apply(lambda x: fuzz.token_set_ratio(input_string, x))

        # Check for first word match and apply additional weight
        first_word_weight = 1.2  # Define a weight for matching the first word
        first_word_scores = series.apply(
            lambda x: first_word_weight if x.lower().split()[0] == input_string.lower().split()[0] else 1
        )

        # Weight the maximum score with the first word match
        max_scores = pd.Series([max(a, b) * w for a, b, w in zip(sort_scores, set_scores, first_word_scores)])

        # Create a dataframe
        df = pd.DataFrame({
            'Value': series,
            'Score': max_scores
        })
        # Sort the dataframe by score in descending order
        df = df.sort_values(by='Score', ascending=False)
        return df

    def display_dataframe_row(self, row_df, target_frame):
        # Clear any previous data
        for widget in target_frame.winfo_children():
            widget.destroy()

        # Display the row data
        for col, value in row_df.items():
            label = tk.Label(target_frame, text=f"{col}: {value.values[0]}")
            label.pack(padx=5, pady=5, anchor='w')

    def append_rows(self, index1, index2, df1, df2, df3):
        # Extract rows from df1 and df2
        row1 = df1.iloc[[index1]].reset_index(drop=True)
        row2 = df2.iloc[[index2]].reset_index(drop=True)
        # Concatenate the rows horizontally
        new_row = pd.concat([row1, row2], axis=1)
        # Append the concatenated row to df3
        df3 = pd.concat([df3, new_row], ignore_index=True)
        return df3

    
    def next_item(self):
        # append selections to dataframe expect for at the start of mapping the first item
        if self.next_item_index>0:
            # Filter df for matches only and append those to final df
            self.temp_subset_df = self.temp_subset_df[self.temp_subset_df['IS_A_MATCH'] == 1]
            print(self.temp_subset_df)
            for index, row in self.temp_subset_df.iterrows():
                self.index_1 = (row['spreadsheet_1_index'])
                self.index_2 = (row['spreadsheet_2_index'])
                self.df_final = self.append_rows(self.index_1, self.index_2, self.spreadsheet1, self.spreadsheet2, self.df_final)
            print(self.df_final)

        # Do this except on final item to map
        if self.next_item_index < self.max_index:
            self.temp_row_df = self.spreadsheet1.iloc[[self.next_item_index]]
            self.current_item_to_map = self.temp_row_df.loc[self.next_item_index, self.column1]
            # Display the item you're matching
            self.current_item_var.set(f'{self.current_item_to_map}')
            # Display the row data in middle_left_frame
            self.display_dataframe_row(self.temp_row_df, self.middle_left_inner_frame)
            self.temp_df = self.fuzzy_logic_dataframe(self.current_item_to_map, self.matching_data_series )
            self.temp_subset_df = self.temp_df.iloc[:50].copy() # Take only the top 50 matches
            self.temp_subset_df['IS_A_MATCH'] = 0 # create match column for later use
            self.temp_subset_df['spreadsheet_1_index'] = self.next_item_index
            self.temp_subset_df['spreadsheet_2_index'] = self.temp_subset_df.index
            self.temp_subset_df.reset_index(drop=True, inplace=True)
            self.display_checkboxes(self.temp_subset_df)
            
            #Update next button
            next_text = "Map Next Item: " + str(self.next_item_index+2)
            self.next_button.config(text=next_text)
            # add one to get to the next index
            self.next_item_index += 1
            #do this just before the last item to map
            if self.next_item_index == self.max_index:
                next_text = "Map Last Item: " + str(self.next_item_index+1)
                self.next_button.config(text=next_text)
        # Do this after final item is mapped and added to the final df
        else:
            # Disable Mapping Button after locking in last item
            self.next_button.config(text="Mapping Done", state=DISABLED, bg="white")
            # Green Save button
            self.save_button.config(bg="green")
            # Clear the middle frames
            if self.middle_right_canvas.winfo_exists():
                self.middle_right_canvas.delete("all")
            if self.middle_left_canvas.winfo_exists():
                self.middle_left_canvas.delete("all")
        
        # Display Mappings Dataframe
        self.display_df(self.df_final)

    # New function to display rows from temp_subset_df with checkboxes in middle_right_frame
    def display_checkboxes(self, subset_df):
        if self.middle_right_canvas:
            self.middle_right_canvas.delete("all") 
        self.middle_right_inner_frame = Frame(self.middle_right_canvas)
        self.middle_right_canvas.create_window((0,0), window=self.middle_right_inner_frame, anchor="nw")
        for widget in self.middle_right_inner_frame.winfo_children():
            widget.destroy()
        # Store variable objects associated with each checkbox
        self.checkbox_vars = {}
        for index, row in subset_df.iterrows():
            value = row['Value']
            is_a_match = row['IS_A_MATCH']
            # Variable for checkbox
            var = tk.IntVar(value=is_a_match)
            self.checkbox_vars[index] = var
            # Checkbutton with command to update 'IS_A_MATCH' column
            checkbox = Checkbutton(self.middle_right_inner_frame, text=value, variable=var, command=lambda i=index: self.update_is_a_match(i))
            checkbox.pack(anchor='w')
        self.middle_right_inner_frame.update_idletasks()
        self.middle_right_canvas.config(scrollregion=self.middle_right_canvas.bbox("all"))

    # Function to update 'IS_A_MATCH' column based on checkbox state
    def update_is_a_match(self, index):
        if self.checkbox_vars[index].get() == 1:
            self.temp_subset_df.at[index, 'IS_A_MATCH'] = 1
        else:
            self.temp_subset_df.at[index, 'IS_A_MATCH'] = 0

    # Display the mapping dataframe on the left top middle canvas
    def display_df(self, df):
        # Ensure the tree is clean
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Get column names and data
        cols = df.columns.tolist()
        data = df.values.tolist()

        # Configure the Treeview columns
        self.tree['columns'] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)  # adjust width as needed

        # Add data to the Treeview
        for row in data:
            self.tree.insert('', 'end', values=row)

    def save_selections(self):
        # Save the selected matches to an Excel file
        if self.df_final is not None and len(self.df_final) > 0:
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if filename:
                self.df_final.to_excel(filename, index=False)
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

root = tk.Tk()
root.state('zoomed')  # To maximize the window
app = Application(root)
root.mainloop()