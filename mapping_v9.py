import pandas as pd
from fuzzywuzzy import fuzz, process
from tkinter import Tk, filedialog, StringVar, END, messagebox, OptionMenu, Button, DISABLED, NORMAL, Listbox, Checkbutton, Label, Frame
from tkinter.ttk import Progressbar
from tqdm import tqdm

def load_spreadsheet(application, spreadsheet_number):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        df = pd.read_excel(filename)
        print(f'Spreadsheet {spreadsheet_number} loaded with shape {df.shape}')
        if spreadsheet_number == 1:
            application.spreadsheet1 = df
            application.update_dropdown(application.dropdown1, df.columns, application.column_button1)
            application.load_button1.config(bg='blue')
            application.dropdown1.config(state=NORMAL)
            application.load_button2.config(state=NORMAL) 
        elif spreadsheet_number == 2:
            application.spreadsheet2 = df
            application.update_dropdown(application.dropdown2, df.columns, application.column_button2)
            application.load_button2.config(bg='blue')
            application.dropdown2.config(state=NORMAL)
            application.match_button.config(state=NORMAL)

def match_data(application, s1, s2, progressbar, threshold=70):
    length = len(s1)
    progressbar["maximum"] = length
    application.matches = [] 
    for i in tqdm(range(length), desc="Matching..."):
        matches = process.extract(s1[i][1], s2, scorer=fuzz.token_sort_ratio, limit=None)
        good_matches = [match for match in matches if match[1] >= threshold]
        application.matches.append(((s1[i][0], s1[i][1]), good_matches))
        progressbar["value"] = i
        progressbar.update()
    application.next_item_index = 0
    application.next_item()

def save_selections(application):
    selected_matches = []
    for key, value in application.selections.items():
        s1_match = key
        for var in value:
            s2_match = var.get()
            if s2_match:
                selected_matches.append((s1_match[0], s1_match[1], s2_match[0], s2_match[1]))

    if selected_matches:
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            match_df = pd.DataFrame(selected_matches, columns=['Code1', 'Name1', 'Code2', 'Name2'])
            match_df.to_excel(filename, index=False)
            messagebox.showinfo("Success", "Matches saved successfully.")
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

        self.load_button1 = Button(master, text="Load Spreadsheet 1", command=lambda: load_spreadsheet(self, 1))
        self.load_button1.pack(fill='x')

        self.load_button2 = Button(master, text="Load Spreadsheet 2", command=lambda: load_spreadsheet(self, 2), state=DISABLED)
        self.load_button2.pack(fill='x')

        self.column_button1 = Button(master, text="Select column from Spreadsheet 1", state=DISABLED)
        self.column_button1.pack(fill='x')

        self.column_button2 = Button(master, text="Select column from Spreadsheet 2", state=DISABLED)
        self.column_button2.pack(fill='x')

        self.match_button = Button(master, text="Match Data", command=self.match_data, state=DISABLED)
        self.match_button.pack(fill='x')

        self.next_button = Button(master, text="Next Item", command=self.next_item, state=DISABLED)
        self.next_button.pack(fill='x')

        self.save_button = Button(master, text="Save Matches", command=lambda: save_selections(self), state=DISABLED)
        self.save_button.pack(fill='x')

        self.close_button = Button(master, text="Close", command=self.close_app)
        self.close_button.pack(fill='x')

        self.progressbar = Progressbar(master, length=500)
        self.progressbar.pack(fill='x')

        self.match_frame = Frame(master)
        self.match_frame.pack(fill='both', expand=True)

        self.matches = []
        self.next_item_index = 0
        self.selections = {}
    
    def close_app(self):
        self.master.destroy()

    def set_column(self, dropdown, value):
        if dropdown == self.dropdown1:
            self.column1 = value
            self.load_button1.config(text=f"Spreadsheet 1: {value}", bg="blue")
        elif dropdown == self.dropdown2:
            self.column2 = value
            self.load_button2.config(text=f"Spreadsheet 2: {value}", bg="blue")
        self.check_if_match_possible()

    def update_dropdown(self, dropdown, options, column_button):
        dropdown['menu'].delete(0, 'end')
        for option in options:
            dropdown['menu'].add_command(label=option, command=lambda value=option: self.set_column(dropdown, value))
        column_button.config(state=NORMAL)

    def match_data(self):
        if self.spreadsheet1 is not None and self.spreadsheet2 is not None and self.column1 is not None and self.column2 is not None:
            s1 = list(zip(self.spreadsheet1['Code'], self.spreadsheet1[self.column1]))
            s2 = list(zip(self.spreadsheet2['Code'], self.spreadsheet2[self.column2]))
            match_data(self, s1, s2, self.progressbar)
            self.match_button.config(bg='blue')
            self.save_button.config(state=NORMAL)
            self.next_button.config(state=NORMAL)

    def next_item(self):
        if self.next_item_index < len(self.matches):
            current_item = self.matches[self.next_item_index]
            self.selections[current_item[0]] = []
            Label(self.match_frame, text=current_item[0][1]).pack()
            for match in current_item[1]:
                match_var = StringVar()
                match_var.set((match[0][0], match[0][1]))
                Checkbutton(self.match_frame, text=f"{match[0][1]} ({match[1]})", variable=match_var, onvalue=(match[0][0], match[0][1]), offvalue="").pack()
                self.selections[current_item[0]].append(match_var)
            self.next_item_index += 1
        else:
            messagebox.showinfo("Done", "All matches have been reviewed.")

    def check_if_match_possible(self):
        if self.spreadsheet1 is not None and self.spreadsheet2 is not None and self.column1 is not None and self.column2 is not None:
            self.match_button.config(state=NORMAL)
        else:
            self.match_button.config(state=DISABLED)


if __name__ == '__main__':
    root = Tk()
    app = Application(root)
    root.mainloop()
