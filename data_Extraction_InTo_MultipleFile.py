import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class ExcelDataTransferApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel Data Transfer")
        self.master.geometry("500x250+450+200")  # Set initial window size 
        self.master.resizable(False, False)  # Disable resizable 

        # Variable to store selected Excel file path
        self.file_path_var = tk.StringVar()

        # Variable to store column names
        self.column_names_var = tk.StringVar()

        #variable to store the folder path
        self.folder_path_var = tk.StringVar()

        # Create and place widgets
        self.create_widgets()

    def create_widgets(self):
        # Label and Entry for selecting Excel file
        file_label = tk.Label(self.master,  font=('Times New Roman', 13, 'bold'),text="Select Excel File:")
        file_label.grid(row=1, column=0, pady=10)

        file_entry = tk.Entry(self.master, textvariable=self.file_path_var, width=30)
        file_entry.grid(row=1, column=1, padx=10, pady=10)

        file_button = tk.Button(self.master, font=('Times New Roman', 12, 'bold'), text="Browse",bd=1, command=self.select_file)
        file_button.grid(row=1, column=2, pady=10)

        # Label and Entry for entering column names
        column_label = tk.Label(self.master, font=('Times New Roman', 13, 'bold'), text="Column Names:")
        column_label.grid(row=2, column=0, pady=10)

        column_entry = tk.Entry(self.master, textvariable=self.column_names_var, width=30)
        column_entry.grid(row=2, column=1, padx=10, pady=10)

        # Label and Entry for the selecting the folder
        file_label2 = tk.Label(self.master , font=('Times New Roman', 13, 'bold'), text="Select the folder:")
        file_label2.grid(row=3 ,column=0, pady=10)

        file_entry2 = tk.Entry(self.master, textvariable=self.folder_path_var, width=30)
        file_entry2.grid(row=3, column=1, padx=10, pady=10)

        file_button2 = tk.Button(self.master, text="Browse", bd=1, font=('Times New Roman', 12, 'bold'),command=self.select_folder)
        file_button2.grid(row=3, column=2, pady=10)

        # Buttons for Generate Files and Transfer Data
        generate_button = tk.Button(self.master, text="Generate Files", bd=1, font=('Times New Roman', 12, 'bold'),
                                    width=30,height=1,command=self.validate_and_generate)
        generate_button.grid(row=4, column=1, pady=20)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        self.file_path_var.set(file_path)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        self.folder_path_var.set(folder_path)
        
    def validate_and_generate(self):
        if self.validate_inputs():
            self.generate_files()

    def validate_inputs(self):
        file_path = self.file_path_var.get()
        column_names = self.column_names_var.get()
        folder_path =self.folder_path_var.get()

        if not file_path:
            messagebox.showerror("Error", "Please select an Excel file.")
            return False
        elif not column_names:
            messagebox.showerror("Error", "Please enter column names.")
            return False
        elif not folder_path:
            messagebox.showerror("Error", "Please select an Folder.")
            return False
        
        return True

    def generate_files(self):

        #Select the column name
        column_names = self.column_names_var.get()

        # Read the Excel file
        file_path = self.file_path_var.get()
        df = pd.read_excel(file_path)

        # Extract unique values from the column name
        unique_values = df[column_names].unique()
        
        # print(unique_values)


        # Create Excel files for each unique value
        for values in unique_values:
            # Filter the DataFrame for the current value
            values_df = df[df[column_names] == values]

            # Create a new Excel file for the current country
            output_file_path = f'{self.folder_path_var.get()}/{values}_data.xlsx'
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                values_df.to_excel(writer, sheet_name='Sheet1', index=False)

        # print("Excel files created successfully.")
        messagebox.showinfo("Success", "Files generated successfully!")

    # def transfer_data(self):
    #     # Add your logic for transferring data here
    #     messagebox.showinfo("Success", "Data transferred successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDataTransferApp(root)
    root.mainloop()
