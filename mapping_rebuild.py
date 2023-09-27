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
                self.update_dropdown(self.dropdown1, df.columns)
                self.dropdown1.config(state=NORMAL)

                # Activate the 'load spreadsheet 2' button
                self.load_button2.config(state=NORMAL)  # This line makes the button active

            elif spreadsheet_number == 2:
                self.spreadsheet2 = df
                self.load_button2.config(bg="blue")
                self.update_dropdown(self.dropdown2, df.columns)
                self.dropdown2.config(state=NORMAL)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading the file:\n{e}")
