import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import pandas as pd
from docx import Document
import sqlite3
from tkinter import END
from tkinter import Toplevel
from openpyxl import load_workbook
import re
import subprocess
import sys
import openpyxl
import webbrowser
from pathlib import Path
from tkcalendar import Calendar
import tkinter.font as tkFont
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment






# Prototyping (make it work, then make it pretty.)

class DatabaseManager: #DB practice(use txt to store folder paths when program finished for faster reads.)

    def __init__(self, db_name='inventory_management.db'):
        self.conn = sqlite3.connect(db_name)
        self.cur = self.conn.cursor()
        self.setup_database()

    def setup_database(self):
        self.cur.execute('''
            CREATE TABLE IF NOT EXISTS folder_paths (
                Folder TEXT PRIMARY KEY,
                Path TEXT
            )
        ''')
        self.conn.commit()

    def save_folder_path(self, folder, path):
        self.cur.execute('''
            INSERT INTO folder_paths (Folder, Path) VALUES (?, ?)
            ON CONFLICT(Folder) DO UPDATE SET Path = excluded.Path;
        ''', (folder, path))
        self.conn.commit()

    def get_folder_path(self, folder_name):
        self.cur.execute('SELECT Path FROM folder_paths WHERE Folder = ?', (folder_name,))
        result = self.cur.fetchone()
        return result[0] if result else None

    def get_all_folders(self):
        self.cur.execute('SELECT Folder FROM folder_paths')
        return [row[0] for row in self.cur.fetchall()]

    def delete_all_folders(self):
        self.cur.execute('DELETE FROM folder_paths')
        self.conn.commit()
        
    def commit_changes(self):
        self.conn.commit()
        
    def __del__(self):
        if hasattr(self, 'conn'):
            self.conn.close()

class ExcelManager:

    def __init__(self, filepath=None, sheet_name=None):
        self.filepath = filepath
        self.sheet_name = sheet_name
        self.data_frame = None

    def load_data(self):
        if self.filepath and self.sheet_name:
            self.data_frame = pd.read_excel(self.filepath, sheet_name=self.sheet_name, engine='openpyxl')
            # Cast all columns to object dtype after loading data
            self.data_frame = self.data_frame.astype('object')

    def get_product_info(self, product_id):
        if self.data_frame is not None:
            # Convert both the product_id and the 'Product ID' column to lower case for comparison
            query_result = self.data_frame[self.data_frame['Product ID'].str.upper() == product_id.upper()]
            if not query_result.empty:
                return query_result.iloc[0].to_dict()
        return None

    def save_product_info(self, product_id, product_data):
        if self.filepath:
            try:
                #print(f"Loading workbook from {self.filepath}")
                workbook = load_workbook(self.filepath)
                #print(f"Accessing sheet {self.sheet_name}")
                sheet = workbook[self.sheet_name]

                # Start by finding the column index for product IDs
                product_id_col_index = self.get_column_index_by_header(sheet, 'Product ID')
                if not product_id_col_index:
                    #print("Product ID column not found")
                    return

                # Update product_data dictionary to convert boolean to YES/NO strings
                for key, value in product_data.items():
                    if isinstance(value, bool):
                        product_data[key] = 'YES' if value else 'NO'

                # Now iterate over the rows to find the matching product ID
                for row in sheet.iter_rows(min_col=product_id_col_index, max_col=product_id_col_index):
                    cell = row[0]
                    if cell.value and str(cell.value).strip().upper() == product_id.upper():
                        row_num = cell.row
                        for key, value in product_data.items():
                            col_index = self.get_column_index_by_header(sheet, key)
                            if col_index:
                                # Special handling for 'To Sell After' date
                                if key == 'To Sell After' and isinstance(value, datetime):
                                    value = value.strftime('%m/%d/%Y')  # Format the date
                                sheet.cell(row=row_num, column=col_index, value=value)
                        workbook.save(self.filepath)
                        break
                else:
                    #print(f"Product ID {product_id} not found in the sheet.")
                    pass
            except Exception as e:
                #print(f"Failed to save changes to Excel file: {e}")
                raise

    @staticmethod
    def get_column_index_by_header(sheet, header_name):
        """
        Gets the column index based on the header name.
        :param sheet: The sheet to search in.
        :param header_name: The header name to find.
        :return: The index of the column, or None if not found.
        """
        for col in sheet.iter_rows(min_row=1, max_row=1, values_only=True):
            if header_name in col:
                return col.index(header_name) + 1
        return None


class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.db_manager = DatabaseManager()
        self.excel_manager = ExcelManager()
        self.edit_mode = False  # Add this line to initialize the edit_mode attribute
        self.inventory_folder = None
        self.sold_folder = None
        self.to_sell_folder = None
        self.pack(fill='both', expand=True)
        
        
        # Make sure you call this before combining and displaying folders
        self.Main_Window_Widgets() 
        
        # Now it's safe to load settings and combine folders since the list widget is created
        self.load_settings()
        self.combine_and_display_folders()
        
        # Call the methods associated with the settings buttons
        self.update_links_in_excel()  # This corresponds to 'Autofill Excel Data(link, asin, tosellafter)'
        self.update_folders_paths()   # This corresponds to 'Update folder names and paths'

    def load_settings(self):
        # Load settings
        try:
            with open("folders_paths.txt", "r") as file:
                lines = file.read().splitlines()
                self.inventory_folder = lines[0]
                self.sold_folder = lines[1]
                self.to_sell_folder = lines[2] if len(lines) > 2 else None
                # ... The rest of your settings loading code ...
        except FileNotFoundError:
            pass
        # Here you could handle the situation if the file is not found, like setting default paths or prompting the user.

    def save_settings(self):
        # This function is called after selecting the source and sold folders
        # Update the table with the new paths
        self.db_manager.cur.execute('''
            UPDATE folder_paths SET Path = ? WHERE Folder = 'Root Folder'
        ''', (self.inventory_folder,))
        self.db_manager.cur.execute('''
            UPDATE folder_paths SET Path = ? WHERE Folder = 'Sold'
        ''', (self.sold_folder,))
        self.db_manager.conn.commit()

    def check_and_update_product_list(self):
        if not self.search_entry.get():  # Check if the search entry is empty
            folder_count = len(next(os.walk(self.inventory_folder))[1])  # Count folders in the directory
            list_count = self.folder_list.size()  # Count items in the Listbox

            if folder_count != list_count:
                self.combine_and_display_folders()  # Update the list items with folder names

            # Schedule this method to be called again after 10000 milliseconds (10 seconds)
            self.after(10000, self.check_and_update_product_list)

    def Main_Window_Widgets(self):
        self.top_frame = tk.Frame(self)
        self.top_frame.pack(fill='x')

        self.settings_button = tk.Button(self.top_frame, text='Settings', command=self.open_settings_window)
        self.settings_button.pack(side='right')

        self.search_frame = tk.Frame(self)
        self.search_frame.pack(fill='x')

        self.search_label = tk.Label(self.search_frame, text="Enter product name here:")
        self.search_label.pack(anchor='w')

        self.search_entry = tk.Entry(self.search_frame, width=30)  # Same width as the Listbox
        self.search_entry.pack(side='left', fill='x', anchor='w')
        self.search_entry.bind('<KeyRelease>', self.search)

        self.bottom_frame = tk.Frame(self)
        self.bottom_frame.pack(fill='both', expand=True)

        self.list_outer_frame = tk.Frame(self.bottom_frame)
        self.list_outer_frame.pack(side='left', fill='y')

        self.list_frame = tk.Frame(self.list_outer_frame)
        self.list_frame.pack(side='left', fill='both', expand=True)

        self.folder_list = tk.Listbox(self.list_frame, width=30)
        self.folder_list.pack(side='left', fill='both', expand=False)
        self.folder_list.bind('<<ListboxSelect>>', self.display_product_details)

        self.list_scrollbar = tk.Scrollbar(self.list_frame)
        self.list_scrollbar.pack(side='right', fill='y')
        self.folder_list.config(yscrollcommand=self.list_scrollbar.set)
        self.list_scrollbar.config(command=self.folder_list.yview)
        
        self.Product_Form()

    def Product_Form(self):

        self.product_frame = tk.Frame(self.bottom_frame, bg='gray')
        self.product_frame.pack(side='right', fill='both', expand=True) #change pack to grid later
        
        self.save_button = tk.Button(self.product_frame, text='Save', command=self.save, state='disabled')
        self.save_button.grid(row=0, column=20, sticky='w', padx=200, pady=0)

        self.edit_button = tk.Button(self.product_frame, text="Edit", command=self.toggle_edit_mode)
        self.edit_button.grid(row=0, column=20, sticky='w', padx=235, pady=0)
        
        
        # Column 0 Widgets
        self.product_id_var = tk.StringVar()
        self.product_id_label = tk.Label(self.product_frame, text='Product ID')
        self.product_id_label.grid(row=0, column=0, sticky='w', padx=0, pady=0)
        self.product_id_entry = tk.Entry(self.product_frame, textvariable=self.product_id_var, state='disabled')
        self.product_id_entry.grid(row=1, column=0, sticky='w', padx=0, pady=0)

        self.product_name_var = tk.StringVar()
        self.product_name_label = tk.Label(self.product_frame, text='Product Name')
        self.product_name_label.grid(row=2, column=0, sticky='w', padx=0, pady=0)
        self.product_name_entry = tk.Entry(self.product_frame, textvariable=self.product_name_var, state='disabled')
        self.product_name_entry.grid(row=3, column=0, sticky='w', padx=0, pady=0)
        
        self.product_folder_var = tk.StringVar()
        self.product_folder_label = tk.Label(self.product_frame, text='Product Folder')
        self.product_folder_label.grid(row=4, column=0, sticky='w', padx=0, pady=0)
        self.product_folder_link = tk.Button(self.product_frame, textvariable=self.product_folder_var, fg="blue", text='No Folder')
        self.product_folder_link.grid(row=5, column=0, sticky='w', padx=0, pady=0)

        self.order_link_var = tk.StringVar()
        self.order_link_label = tk.Label(self.product_frame, text='Order Link')
        self.order_link_label.grid(row=6, column=0, sticky='w', padx=0, pady=0)
        
        # Replace the Entry with a Text widget for clickable links
        self.order_link_text = tk.Text(self.product_frame, height=1, width=30, font="TkDefaultFont")
        self.order_link_text.grid(row=7, column=0, sticky='w', padx=0, pady=0)
        self.order_link_text.tag_configure("hyperlink", foreground="blue", underline=True)
        self.order_link_text.bind("<Button-1>", self.open_hyperlink)

        self.asin_var = tk.StringVar()
        self.asin_label = tk.Label(self.product_frame, text='ASIN')
        self.asin_label.grid(row=8, column=0, sticky='w', padx=0, pady=0)
        self.asin_entry = tk.Entry(self.product_frame, textvariable=self.asin_var, state='disabled')
        self.asin_entry.grid(row=9, column=0, sticky='w', padx=0, pady=0)

        # Column 4 Widgets
        # Assuming you want to create a spacer between column 0 and column 1
        self.product_frame.grid_columnconfigure(2, minsize=20)  # This creates a 20-pixel-wide empty column as spacer
        
        self.order_date_var = tk.StringVar()
        self.order_date_label = tk.Label(self.product_frame, text='Order Date')
        self.order_date_label.grid(row=0, column=4, sticky='w', padx=0, pady=0)
        self.order_date_entry = tk.Entry(self.product_frame, textvariable=self.order_date_var, state='disabled')
        self.order_date_entry.grid(row=1, column=4, sticky='w', padx=0, pady=0)

        self.to_sell_after_var = tk.StringVar()
        self.to_sell_after_label = tk.Label(self.product_frame, text='To Sell After')
        self.to_sell_after_label.grid(row=2, column=4, sticky='w', padx=0, pady=0)
        self.to_sell_after_entry = tk.Entry(self.product_frame, textvariable=self.to_sell_after_var, state='disabled')
        self.to_sell_after_entry.grid(row=3, column=4, sticky='w', padx=0, pady=0)

        # Column 8 Widgets
        self.product_frame.grid_columnconfigure(6, minsize=20)  # This creates a 20-pixel-wide empty column as spacer
        self.sold_var = tk.BooleanVar()
        self.sold_checkbutton = tk.Checkbutton(self.product_frame, text='Sold', variable=self.sold_var, state='disabled')
        self.sold_checkbutton.grid(row=0, column=8, sticky='w', padx=0, pady=0)

        self.sold_date_var = tk.StringVar()
        self.sold_date_label = tk.Label(self.product_frame, text='Sold Date')
        self.sold_date_label.grid(row=1, column=8, sticky='w', padx=0, pady=0)
        
        self.sold_date_entry = tk.Entry(self.product_frame, textvariable=self.sold_date_var, state='disabled')
        self.sold_date_entry.grid(row=2, column=8, sticky='w', padx=0, pady=0)

        small_font = tkFont.Font(size=5)  # You can adjust the size as needed
        self.sold_date_button = tk.Button(self.product_frame, text="Pick \nDate", command=self.pick_date, state='disabled', font=small_font)
        self.sold_date_button.grid(row=2, column=8, sticky='e', padx=0, pady=0)

        self.fair_market_value_var = tk.StringVar()
        self.fair_market_value_label = tk.Label(self.product_frame, text='Fair Market Value')
        self.fair_market_value_label.grid(row=3, column=8, sticky='w', padx=0, pady=0)
        self.fair_market_value_entry = tk.Entry(self.product_frame, textvariable=self.fair_market_value_var, state='disabled')
        self.fair_market_value_entry.grid(row=4, column=8, sticky='w', padx=0, pady=0)
        
        self.sold_price_var = tk.StringVar()
        self.sold_price_label = tk.Label(self.product_frame, text='Sold Price')
        self.sold_price_label.grid(row=5, column=8, sticky='w', padx=0, pady=0)
        self.sold_price_entry = tk.Entry(self.product_frame, textvariable=self.sold_price_var, state='disabled')
        self.sold_price_entry.grid(row=6, column=8, sticky='w', padx=0, pady=0)
        
        self.payment_type_var = tk.StringVar()
        self.payment_type_label = tk.Label(self.product_frame, text='Payment Type')
        self.payment_type_label.grid(row=7, column=8, sticky='w', padx=0, pady=0)
        
        self.payment_type_combobox = ttk.Combobox(self.product_frame, textvariable=self.payment_type_var, state='disabled')
        self.payment_type_combobox['values'] = ('', 'Cash', 'ATH Movil')
        self.payment_type_combobox.grid(row=8, column=8, sticky='w', padx=0, pady=0)
        
        # Column 12 Widgets
        self.product_frame.grid_columnconfigure(10, minsize=20)  # This creates a 20-pixel-wide empty column as spacer
        self.cancelled_order_var = tk.BooleanVar()
        self.cancelled_order_checkbutton = tk.Checkbutton(self.product_frame, text='Cancelled Order', variable=self.cancelled_order_var, state='disabled')
        self.cancelled_order_checkbutton.grid(row=0, column=12, sticky='w', padx=0, pady=0)

        self.damaged_var = tk.BooleanVar()
        self.damaged_checkbutton = tk.Checkbutton(self.product_frame, text='Damaged', variable=self.damaged_var, state='disabled')
        self.damaged_checkbutton.grid(row=1, column=12, sticky='w', padx=0, pady=0)

        self.personal_var = tk.BooleanVar()
        self.personal_checkbutton = tk.Checkbutton(self.product_frame, text='Personal', variable=self.personal_var, state='disabled')
        self.personal_checkbutton.grid(row=2, column=12, sticky='w', padx=0, pady=0)

        self.reviewed_var = tk.BooleanVar()
        self.reviewed_checkbutton = tk.Checkbutton(self.product_frame, text='Reviewed', variable=self.reviewed_var, state='disabled')
        self.reviewed_checkbutton.grid(row=4, column=12, sticky='w', padx=0, pady=0)

        self.pictures_downloaded_var = tk.BooleanVar()
        self.pictures_downloaded_checkbutton = tk.Checkbutton(self.product_frame, text='Pictures Downloaded', variable=self.pictures_downloaded_var, state='disabled')
        self.pictures_downloaded_checkbutton.grid(row=5, column=12, sticky='w', padx=0, pady=0)

        # Load settings
        try:
            with open("folders_paths.txt", "r") as file:
                lines = file.read().splitlines()
                self.inventory_folder = lines[0]
                self.sold_folder = lines[1]
                self.to_sell_folder = lines[2] if len(lines) > 2 else None
                if self.inventory_folder:  # Check if inventory_folder is defined
                    self.combine_and_display_folders()
        except FileNotFoundError:
            pass
        self.search_entry.focus_set()

    def pick_date(self):
        def grab_date():
            selected_date = cal.selection_get()  # Get the selected date
            formatted_date = selected_date.strftime('%m/%d/%Y')  # Format the date
            self.sold_date_entry.delete(0, tk.END)  # Clear the entry field
            self.sold_date_entry.insert(0, formatted_date)  # Insert the formatted date
            top.destroy()  # Close the Toplevel window
        def select_today_and_close(event):
            cal.selection_set(datetime.today())  # Set selection to today's date
            grab_date()  # Then grab the date and close
        top = tk.Toplevel(self)
        today = datetime.today()
        cal = Calendar(top, selectmode='day', year=today.year, month=today.month, day=today.day)
        cal.pack(pady=20)    # Set focus to the Toplevel window and bind the Enter key
        top.focus_set()
        top.bind('<Return>', select_today_and_close)
        cal.bind("<<CalendarSelected>>", lambda event: grab_date())

    def focus_search_entry(self):
        self.search_entry.focus_set()

    def open_hyperlink(self, event):
        try:
            start_index = self.order_link_text.index("@%s,%s" % (event.x, event.y))
            tag_indices = list(self.order_link_text.tag_ranges('hyperlink'))
            for start, end in zip(tag_indices[0::2], tag_indices[1::2]):
                if self.order_link_text.compare(start_index, ">=", start) and self.order_link_text.compare(start_index, "<=", end):
                    url = self.order_link_text.get(start, end)
                    webbrowser.open(url)
                    return "break"
        except Exception as e:
            print(f"Error when opening hyperlink: {e}")

    def open_settings_window(self):
        if hasattr(self, 'settings_window') and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return
        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("Settings")
        
        self.settings_window.state('zoomed')
        
        # Bind the closing event to the on_settings_close function
        self.settings_window.protocol("WM_DELETE_WINDOW", self.on_settings_close)

        self.settings_window.update()  # This updates the window and calculates sizes
        window_width = self.settings_window.winfo_width()  # Gets the width of the window
        # Load settings
        self.default_filepath, self.default_sheet = self.load_excel_settings()
        
        # Configure the grid columns 
        self.settings_window.grid_columnconfigure(3, weight=1)

        self.inventory_folder_frame = tk.Frame(self.settings_window)
        self.inventory_folder_frame.grid(row=0, column=0, sticky='w')
        
        self.inventory_folder_button = tk.Button(self.inventory_folder_frame, text="Choose Inventory Folder", command=self.choose_inventory_folder)
        self.inventory_folder_button.grid(row=1, column=0, padx=(window_width//4, 0))  # Half the remaining space to the left
        self.inventory_folder_label = tk.Label(self.inventory_folder_frame, text=self.inventory_folder if self.inventory_folder else "Not chosen")
        self.inventory_folder_label.grid(row=1, column=1, padx=(0, window_width//4), sticky='ew')  # Half the remaining space to the right

        self.sold_folder_frame = tk.Frame(self.settings_window)
        self.sold_folder_frame.grid(row=1, column=0, sticky='w')
        
        self.sold_folder_button = tk.Button(self.sold_folder_frame, text="Choose Sold Inventory Folder", command=self.choose_sold_folder)
        self.sold_folder_button.grid(row=2, column=0, padx=(window_width//4, 0))
        self.sold_folder_label = tk.Label(self.sold_folder_frame, text=self.sold_folder  if self.sold_folder else "Not chosen")
        self.sold_folder_label.grid(row=2, column=1, padx=(0, window_width//4), sticky='ew')

        # Inside the open_settings method of Application class after existing setup code for other buttons

        # Choose Folder with Products to Sell Button and Label
        self.to_sell_folder_frame = tk.Frame(self.settings_window)
        self.to_sell_folder_frame.grid(row=2, column=0, sticky='w')  # Adjust the row index as needed
        
        self.to_sell_folder_button = tk.Button(self.to_sell_folder_frame, text="Choose Products to Sell Folder", command=self.choose_to_sell_folder)
        self.to_sell_folder_button.grid(row=3, column=0, padx=(window_width//4, 0))
        self.to_sell_folder_label = tk.Label(self.to_sell_folder_frame, text=self.to_sell_folder if self.to_sell_folder else "Not chosen")
        self.to_sell_folder_label.grid(row=3, column=1, padx=(0, window_width//4), sticky='ew')

        # Add a new frame for the Excel database selection

        # Excel Database Selection frame and button
        self.excel_db_frame = tk.Frame(self.settings_window)
        self.excel_db_frame.grid(row=3, column=0, sticky='w')  # Adjust row as needed
        
        self.excel_db_button = tk.Button(self.excel_db_frame, text="Select Excel Database", command=self.select_excel_database)
        self.excel_db_button.grid(row=4, column=0, padx=(window_width//4, 0))
        
        excel_db_text = f"{self.default_filepath} - Sheet: {self.default_sheet}" if self.default_filepath and self.default_sheet else "Not chosen"
        self.excel_db_label = tk.Label(self.excel_db_frame, text=excel_db_text)
        self.excel_db_label.grid(row=4, column=1, padx=(0, window_width//4), sticky='ew')
        
        # Add a new button for creating new word documents
        self.create_word_files_button = tk.Button(self.settings_window, text="Create Word Files for Products", command=self.correlate_data)
        self.create_word_files_button.grid(row=5, column=0, padx=(window_width//4, 0), sticky='w')
        
        self.autofill_links_asin_tosellafter_data_button = tk.Button(self.settings_window, text="Autofill Excel Data(link, asin, tosellafter)", command=self.update_links_in_excel)
        self.autofill_links_asin_tosellafter_data_button.grid(row=6, column=0, padx=(window_width//4, 0), sticky='w')
        
        self.update_foldersnames_folderpaths_button = tk.Button(self.settings_window, text="Update folder names and paths", command=self.update_folders_paths)
        self.update_foldersnames_folderpaths_button.grid(row=7, column=0, padx=(window_width//4, 0), sticky='w')
        
        self.products_to_sell_list_button = tk.Button(self.settings_window, text="Show list of products available to sell", command=self.products_to_sell_report)
        self.products_to_sell_list_button.grid(row=8, column=0, padx=(window_width//4, 0), sticky='w')


        self.back_button = tk.Button(self.settings_window, text="<- Back", command=self.back_to_main)
        self.back_button.grid(row=0, column=0, sticky='nw')  # Change this line to place the back button in the fourth row
        
        self.combine_and_display_folders()
        self.settings_window.protocol("WM_DELETE_WINDOW", lambda: on_close(self, self.master))

        self.master.withdraw()
        
    def on_settings_close(self):
        self.master.destroy()
    
    def products_to_sell_report(self):
        # Ensure the Excel file path and sheet name are set
        filepath, sheet_name = self.load_excel_settings()
        if not filepath or not sheet_name:
            messagebox.showerror("Error", "Excel file path or sheet name is not set.")
            return

        # Define the Backup folder path (next to the inventory folder)
        backup_folder = Path(self.inventory_folder).parent / "Excel Backups"
        backup_folder.mkdir(exist_ok=True)

        # Create a copy of the Excel file in the backup folder
        today_str = datetime.now().strftime("%Y-%m-%d")
        copy_path = backup_folder / f"Products To Sell - {today_str}.xlsx"
        shutil.copy2(filepath, copy_path)

        # Load the workbook and get the sheet
        workbook = load_workbook(copy_path)
        original_sheet = workbook[sheet_name]

        # Create a DataFrame from the sheet data
        data = original_sheet.values
        columns = next(data)[0:]
        df = pd.DataFrame(data, columns=columns)

        # Keep only necessary columns
        df = df[['Product ID', 'To Sell After', 'Product Name', 'Fair Market Value']]

        # Convert 'To Sell After' column to datetime and filter rows
        df['To Sell After'] = pd.to_datetime(df['To Sell After'], errors='coerce')
        today = pd.to_datetime('today').normalize()
        df = df.dropna(subset=['To Sell After'])  # Remove rows with empty 'To Sell After'
        filtered_df = df[df['To Sell After'] <= today]

        # Sort the filtered DataFrame by 'To Sell After' in descending order
        sorted_df = filtered_df.sort_values(by='To Sell After', ascending=False)

        # Delete the original sheet and create a new one
        del workbook[sheet_name]
        new_sheet = workbook.create_sheet(sheet_name)

        # Write the sorted DataFrame back to the new sheet
        for r_idx, row in enumerate(dataframe_to_rows(sorted_df, index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                cell = new_sheet.cell(row=r_idx, column=c_idx, value=value)
                # Apply date format to 'To Sell After' column (assuming it's the second column)
                if c_idx == 2 and r_idx > 1:  # Skip header row
                    cell.number_format = 'MM/DD/YYYY'
                # Apply currency format to 'Fair Market Value' column (assuming it's the fourth column)
                if c_idx == 4 and r_idx > 1:  # Skip header row
                    cell.number_format = '"$"#,##0.00'
                # Apply middle and center alignment to all cells
                if c_idx == 3:  # 'Product Name' column
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                else:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        # Define the table dimensions
        table_ref = f"A1:{chr(65 + sorted_df.shape[1] - 1)}{sorted_df.shape[0] + 1}"

        # Create a table
        table = Table(displayName="ProductsToSellTable", ref=table_ref)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                            showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        table.tableStyleInfo = style
        new_sheet.add_table(table)

        # Adjust column widths
        new_sheet.column_dimensions['A'].width = 120 / 7  # Width for 'Product ID'
        new_sheet.column_dimensions['B'].width = 120 / 7  # Width for 'To Sell After'
        new_sheet.column_dimensions['C'].width = 700 / 7  # Width for 'Product Name'
        new_sheet.column_dimensions['D'].width = 120 / 7  # Width for 'Fair Market Value'

        # Save the changes to the workbook
        workbook.save(copy_path)


        # Open the modified Excel file
        if sys.platform == "win32":
            os.startfile(copy_path)
        elif sys.platform == "darwin":  # macOS
            subprocess.run(["open", copy_path])
        else:  # Linux variants
            subprocess.run(["xdg-open", copy_path])

    def exit_correlate_window(self):
        self.correlate_window.destroy()
        self.open_settings_window()

    def back_to_main(self):
        self.settings_window.destroy()
        self.master.deiconify()
        self.master.state('zoomed')
        
        # Load settings again in case they were changed
        self.load_settings()
        
        # Refresh the folder list with the updated settings
        self.combine_and_display_folders()
        
        self.focus_search_entry()

    def choose_inventory_folder(self):
        inventory_folder = filedialog.askdirectory()
        if inventory_folder:
            self.inventory_folder = inventory_folder
            self.inventory_folder_label.config(text=inventory_folder)  # Update the label directly
            self.save_settings()
            self.combine_and_display_folders()

    @staticmethod
    def custom_sort_key(s):
        # A regular expression to match words in the folder name.
        # Words are defined as sequences of alphanumeric characters and underscores.
        words = re.findall(r'\w+', s.lower())
        
        # The key will be a tuple consisting of the length of the first word,
        # the first word itself (for alphanumeric sorting), and then the rest of the words.
        # Lowercase all words for case-insensitive comparison, numbers will sort naturally before letters.
        return (len(words[0]),) + tuple(words)

    def combine_and_display_folders(self):
        # Clear the folder list first
        self.folder_list.delete(0, tk.END)

        # Begin a transaction
        self.db_manager.cur.execute("BEGIN")
        try:
            # Combine the folders from all three paths
            combined_folders = []
            for folder_path in [self.inventory_folder, self.sold_folder, self.to_sell_folder]:
                if folder_path and os.path.exists(folder_path):
                    for root, dirs, files in os.walk(folder_path):
                        for dir_name in dirs:
                            combined_folders.append(dir_name)
                            full_path = os.path.join(root, dir_name)
                            # Update the database with the current folder paths
                            self.db_manager.cur.execute("INSERT OR REPLACE INTO folder_paths (Folder, Path) VALUES (?, ?)", (dir_name, full_path))
            self.db_manager.conn.commit()  # Commit the transaction if all is well
        except Exception as e:
            self.db_manager.conn.rollback()  # Rollback if there was an error
            messagebox.showerror("Database Error", f"An error occurred while updating the folder paths: {e}")
        # Deduplicate folder names
        unique_folders = list(set(combined_folders))

        # Sort using the custom sort key function
        sorted_folders = sorted(unique_folders, key=self.custom_sort_key)

        # Insert the sorted folders into the list widget
        for folder in sorted_folders:
            self.folder_list.insert(tk.END, folder)

    def choose_sold_folder(self):
        self.sold_folder = filedialog.askdirectory()
        if self.sold_folder:
            self.sold_folder_label.config(text=self.sold_folder)  # Update the label directly
            self.save_settings()
        # Update the Sold Folder path
        self.db_manager.cur.execute('''
            INSERT INTO folder_paths (Folder, Path) VALUES ('Sold', ?)
            ON CONFLICT(Folder) DO UPDATE SET Path = excluded.Path;
        ''', (self.sold_folder,))

        self.db_manager.conn.commit()

    def choose_to_sell_folder(self):
        self.to_sell_folder = filedialog.askdirectory()
        if self.to_sell_folder:
            self.to_sell_folder_label.config(text=self.to_sell_folder)
            self.save_settings()  # Save the settings including the new folder path

    def save_settings(self):
        # Here you will gather all the paths and write them to the settings.txt file
        with open("folders_paths.txt", "w") as file:
            file.write(f"{self.inventory_folder}\n{self.sold_folder}\n{self.to_sell_folder}")

    def search(self, event):
        search_terms = self.search_entry.get().split()  # Split the search string into words
        if search_terms:
            self.folder_list.delete(0, END)  # Clear the current list

            # Walk the directory tree from the inventory_folder path
            for root, dirs, files in os.walk(self.inventory_folder):
                # Check if 'dirs' is empty, meaning 'root' is a leaf directory
                if not dirs:
                    folder_name = os.path.basename(root)  # Get the name of the leaf directory
                    # Check if all search terms are in the folder name (case insensitive)
                    if all(term.upper() in folder_name.upper() for term in search_terms):
                        self.folder_list.insert(END, folder_name)
        else:
            self.combine_and_display_folders()  # If the search box is empty, display all folders

    def display_product_details(self, event):

        selection = self.folder_list.curselection()
        # Get the index of the selected item
        if not selection:
            return  # No item selected
        index = selection[0]
        selected_folder_name = self.folder_list.get(index)
        selected_product_id = selected_folder_name.split(' ')[0].upper()  # Assuming the product ID is at the beginning
        if self.edit_mode:
            self.toggle_edit_mode()
        # Ensure that the Excel file path and sheet name are set
        filepath, sheet_name = self.load_excel_settings()
        if filepath and sheet_name:
            self.excel_manager.filepath = filepath
            self.excel_manager.sheet_name = sheet_name
            self.excel_manager.load_data()  # Load the data

            # Retrieve product information from the DataFrame
            try:
                product_info = self.excel_manager.get_product_info(selected_product_id)
                # Right after fetching product_info
                self.product_folder_path = self.get_folder_path_from_db(selected_product_id)

                if product_info:

                    self.cancelled_order_var.set(self.excel_value_to_bool(product_info.get('Cancelled Order')))
                    self.damaged_var.set(self.excel_value_to_bool(product_info.get('Damaged')))
                    self.personal_var.set(self.excel_value_to_bool(product_info.get('Personal')))
                    self.reviewed_var.set(self.excel_value_to_bool(product_info.get('Reviewed')))
                    self.pictures_downloaded_var.set(self.excel_value_to_bool(product_info.get('Pictures Downloaded')))
                    self.sold_var.set(self.excel_value_to_bool(product_info.get('Sold')))
                    
                    # For each field, check if the value is NaN using pd.isnull and set it to an empty string if it is
                    self.asin_var.set('' if pd.isnull(product_info.get('ASIN')) else product_info.get('ASIN', ''))
                    self.product_id_var.set('' if pd.isnull(product_info.get('Product ID')) else product_info.get('Product ID', ''))
                    # ... handle the product image ...
                    self.product_name_var.set('' if pd.isnull(product_info.get('Product Name')) else product_info.get('Product Name', ''))
                    
                    # When a product is selected and the order date is fetched
                    order_date = product_info.get('Order Date', '')
                    formatted_order_date = ''  # Default value
                    if isinstance(order_date, datetime):
                        formatted_order_date = order_date.strftime('%m/%d/%Y')
                        self.order_date_var.set(formatted_order_date)
                    elif isinstance(order_date, str) and order_date:
                        try:
                            # If the date is in the format 'mm/dd/yy', such as '2/15/23'
                            order_date = datetime.strptime(order_date, "%m/%d/%Y")
                            formatted_order_date = order_date.strftime('%m/%d/%Y')
                            self.order_date_var.set(formatted_order_date)
                        except ValueError as e:
                            messagebox.showerror("Error", f"Incorrect date format: {e}")
                    else:
                        self.order_date_var.set('')
                    self.order_date_var.set(formatted_order_date)
                    
                    # When a product is selected and the order date is fetched
                    to_sell_after = product_info.get('To Sell After', '')
                    formatted_to_sell_after = ''  # Default value
                    if pd.notnull(to_sell_after):  # Check if 'To Sell After' is not null
                        try:
                            if isinstance(to_sell_after, datetime):
                                formatted_to_sell_after = to_sell_after.strftime('%m/%d/%Y')
                            elif isinstance(to_sell_after, str) and to_sell_after:
                                to_sell_after = datetime.strptime(to_sell_after, "%m/%d/%Y")
                                formatted_to_sell_after = to_sell_after.strftime('%m/%d/%Y')
                        except ValueError as e:
                            messagebox.showerror("Error", f"Incorrect date format: {e}")
                    self.to_sell_after_var.set(formatted_to_sell_after)
                    self.update_to_sell_after_color()
                    
                    sold_date = product_info.get('Sold Date', '')
                    formatted_sold_date = ''  # Default value

                    if pd.notnull(sold_date):  # Check if 'Sold Date' is not null
                        try:
                            if isinstance(sold_date, datetime):
                                formatted_sold_date = sold_date.strftime('%m/%d/%Y')
                            elif isinstance(sold_date, str) and sold_date:
                                # Parse the date string to a date object and format it
                                sold_date = datetime.strptime(sold_date, "%m/%d/%Y").date()
                                formatted_sold_date = sold_date.strftime('%m/%d/%Y')
                        except ValueError as e:
                            messagebox.showerror("Error", f"Incorrect date format: {e}")
                    else:
                        formatted_sold_date = ''

                    self.sold_date_var.set(formatted_sold_date)

                    self.fair_market_value_var.set('' if pd.isnull(product_info.get('Fair Market Value')) else product_info.get('Fair Market Value', ''))
                    self.order_link_text.delete(1.0, "end")
                    hyperlink = product_info.get('Order Link', '')
                    if hyperlink:
                        self.order_link_text.insert("insert", hyperlink, "hyperlink")
                        self.order_link_text.tag_add("hyperlink", "1.0", "end")
                    self.sold_price_var.set('' if pd.isnull(product_info.get('Sold Price')) else product_info.get('Sold Price', ''))
                    self.payment_type_var.set('' if pd.isnull(product_info.get('Payment Type')) else product_info.get('Payment Type', ''))
                    # ... continue with other fields as needed ...
                    # Add code here to populate the Sold Date and other date-related fields, if applicable
                    
                    # Fetch the full folder path from the database using the product ID.
                    folder_path = self.get_folder_path_from_db(selected_product_id)

                    # Extract the name of the parent directory (where the product folder is located)
                    parent_folder_name = os.path.basename(os.path.dirname(folder_path)) if folder_path else "No Folder"
                    self.product_folder_var.set(parent_folder_name)

                    # If the folder path exists, update the button to open the product folder when clicked
                    if folder_path and os.path.exists(folder_path):
                        self.product_folder_link.config(command=lambda: self.open_product_folder(folder_path), state='normal')
                    else:
                        self.product_folder_var.set("No Folder")
                        self.product_folder_link.config(state='disabled')
                    
                else:
                    self.cancelled_order_var.set(False)
                    self.damaged_var.set(False)
                    self.personal_var.set(False)
                    self.reviewed_var.set(False)
                    self.pictures_downloaded_var.set(False)
                    self.sold_var.set(False)
                    
                    # Populate the widgets with the matched data
                    self.asin_var.set('')
                    self.product_id_var.set('')
                    self.to_sell_after_var.set('')
                    # Add code here to handle the product image, if applicable
                    self.product_name_var.set('Product not found in Excel.')
                    self.order_date_var.set('')
                    self.fair_market_value_var.set('')
                    self.order_link_text.delete(1.0, "end")
                    self.sold_price_var.set('')
                    self.payment_type_var.set('')
                    self.sold_date_var.set('')
                    self.product_folder_var.set("No Folder")
                    self.product_folder_link.config(state='disabled')

            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")
                #print(f"Error retrieving product details: {e}")
        else:
            messagebox.showerror("Error", "Excel file path or sheet name is not set.")

        # Unbind the Enter key from the save_button's command
        self.master.unbind('<Return>')
        
        # Bind the Escape key to do nothing, which overrides the binding if in edit mode
        self.master.bind('<Escape>', lambda e: None)
        # Any other code you want to execute when displaying product details, such as configuring widget states
        
        # Now bind the Enter key to the edit_button's command
        self.edit_button.focus_set()  # Optional: set the focus on the edit button
        self.master.bind('<Return>', lambda e: self.edit_button.invoke())

    def open_product_folder(self, folder_path):
        if sys.platform == "win32":
            os.startfile(folder_path)
        elif sys.platform == "darwin":  # macOS
            subprocess.run(["open", folder_path])
        else:  # Linux variants
            subprocess.run(["xdg-open", folder_path])

    def excel_value_to_bool(self, value):
        # Check for NaN explicitly and return False if found
        if pd.isnull(value):
            return False
        if isinstance(value, str):
            return value.strip().lower() in ['yes', 'true', '1']
        elif isinstance(value, (int, float)):
            return bool(value)
        return False

    def update_to_sell_after_color(self):
        # Get today's date
        today = date.today()

        # Get the date from the to_sell_after_var entry
        to_sell_after_str = self.to_sell_after_var.get()
        if to_sell_after_str:
            try:
                # Parse the date string to a date object
                to_sell_after_date = datetime.strptime(to_sell_after_str, "%m/%d/%Y").date()

                # If the to_sell_after date is today or has passed, change the label's background color to green
                if to_sell_after_date <= today:
                    self.to_sell_after_label.config(background='green')
                else:
                    self.to_sell_after_label.config(background='white')
            except ValueError:
                # If there's a ValueError, it means the string was not in the expected format
                # Handle incorrect date format or clear the background
                self.to_sell_after_label.config(background='white')

    def toggle_edit_mode(self):
        # Toggle the edit mode
        print("toggling edit mode")
        
        # Toggle the edit mode state. Switch self.edit_mode to True if it's currently False, and vice versa. 
        # This line essentially switches between edit and view modes for the application.
        self.edit_mode = not self.edit_mode
        # Set the state variable depending on self.edit_mode. 
        # When in edit mode (self.edit_mode is True), state is set to 'normal', enabling interaction and editing of widgets. 
        # When not in edit mode (self.edit_mode is False), state is set to 'disabled', making widgets non-interactive and uneditable.
        state = 'normal' if self.edit_mode else 'disabled' 
        # Set the readonly_state variable based on self.edit_mode. 
        # Use 'readonly' for specific widgets when in edit mode to allow viewing but restrict modification. 
        # Set to 'disabled' when not in edit mode to make these widgets non-interactive.
        readonly_state = 'readonly' if self.edit_mode else 'disabled'  # Use 'readonly' when in edit mode, 'disabled' otherwise
        
        self.sold_checkbutton.config(state='disabled')
        self.cancelled_order_checkbutton.config(state=state)
        self.damaged_checkbutton.config(state=state)
        self.personal_checkbutton.config(state=state)
        self.reviewed_checkbutton.config(state=state)
        self.pictures_downloaded_checkbutton.config(state=state)
        self.order_date_entry.config(state='disabled')
        self.sold_date_entry.config(state=state)
        self.sold_date_button.config(state=state)
        self.to_sell_after_entry.config(state='disabled')
        self.payment_type_combobox.config(state=readonly_state)
        self.asin_entry.config(state=state)
        self.product_id_entry.config(state='disabled')
        self.product_name_entry.config(state='disabled')
        self.fair_market_value_entry.config(state=state)
        self.sold_price_entry.config(state=state)
        self.save_button.config(state=state)
        if self.edit_mode:

            self.save_button.focus_set()  # Optional: set the focus on the save button
            # When in edit mode, bind the Enter key to the save_button's command
            self.master.bind('<Return>', lambda e: self.save_button.invoke())
            
            # When in edit mode, bind the Escape key to the edit_button's command
            self.master.bind('<Escape>', lambda e: self.edit_button.invoke())
        else:
            # When not in edit mode, unbind the Enter and Escape keys
            self.master.unbind('<Return>')
            self.master.unbind('<Escape>')
            self.edit_button.focus_set()  # Optional: set the focus on the save button
            # When in edit mode, bind the Enter key to the save_button's command
            self.master.bind('<Return>', lambda e: self.edit_button.invoke())

    def save(self):
        # Update the 'Sold' checkbox based on the 'Sold Date' entry
        if self.sold_date_var.get():
            # If 'Sold Date' is not empty, check 'Sold'
            self.sold_var.set(True)
        else:
            # If 'Sold Date' is empty, uncheck 'Sold'
            self.sold_var.set(False)
        product_id = self.product_id_var.get().strip().upper()

        # Ensure that the Excel file path and sheet name are set.
        filepath, sheet_name = self.load_excel_settings()

        if not filepath or not sheet_name:
            messagebox.showerror("Error", "Excel file path or sheet name is not set.")
            return
            
        # Collect the data from the form.
        product_data = {
            'Cancelled Order': self.cancelled_order_var.get(),
            'Damaged': self.damaged_var.get(),
            'Personal': self.personal_var.get(),
            'Reviewed': self.reviewed_var.get(),
            'Pictures Downloaded': self.pictures_downloaded_var.get(),
            'Sold': self.sold_var.get(),
            'To Sell After': self.to_sell_after_var.get(),
            'Product Name': self.product_name_var.get(),
            'Fair Market Value': self.fair_market_value_var.get(),
            'Sold Price': self.sold_price_var.get(),
            'Payment Type': self.payment_type_var.get(),
            'Sold Date': self.sold_date_var.get(),
            # ... and so on for the rest of your form fields.
        }

        # Use the ExcelManager method to save the data.
        try:
            self.excel_manager.save_product_info(product_id, product_data)
            messagebox.showinfo("Success", "Product information updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save changes to Excel file: {e}")
            return
        
        # Folder movement logic
        current_folder_path = self.get_folder_path_from_db(product_id)
        #print(f"Current folder path: {current_folder_path}")

        if not current_folder_path:
            messagebox.showerror("Error", f"No current folder path found for Product ID {product_id}")
            return
        
        target_folder_path = None

        # Decide target folder based on the sold status and to sell after date
        if self.sold_var.get():
            target_folder_path = self.sold_folder
        else:
            to_sell_after_str = self.to_sell_after_var.get()
            if to_sell_after_str:
                try:
                    to_sell_after_date = datetime.strptime(to_sell_after_str, "%m/%d/%Y").date()
                    today = date.today()
                    if to_sell_after_date <= today:
                        target_folder_path = self.to_sell_folder
                    else:
                        target_folder_path = self.inventory_folder
                except ValueError as e:
                    messagebox.showerror("Error", f"Invalid 'To Sell After' date format: {e}")
                    return
        # Use #print statements to debug the current and target folder paths
        #print(f"Current folder path: {current_folder_path}")
        #print(f"Target folder path: {target_folder_path}")
        # Check if the target folder is determined and it's not the same as the current folder
        if target_folder_path and os.path.isdir(current_folder_path) and current_folder_path != target_folder_path:
            try:
                # Perform the move operation
                new_folder_path = shutil.move(current_folder_path, target_folder_path)
                
                # Save the new folder path in the database
                self.db_manager.save_folder_path(product_id, new_folder_path)
                
                # Refresh the product folder path attribute to the new path
                self.product_folder_path = new_folder_path
                
                # Ensure that changes are committed to the database
                self.db_manager.commit_changes()
                
                messagebox.showinfo("Folder Moved", f"Folder for '{product_id}' moved successfully to the new location.")
                #print(f"Folder for '{product_id}' moved from {current_folder_path} to {new_folder_path}")
                self.refresh_and_select_product(product_id)


            except Exception as e:
                messagebox.showerror("Error", f"Failed to move the folder: {e}")
                self.refresh_and_select_product(product_id)
                
        doc_data = (product_id, product_id, self.product_name_var.get())  # Construct the doc_data tuple
        self.create_word_doc(doc_data, iid="dummy", show_message=True)  # Call create_word_doc with dummy iid

        self.toggle_edit_mode()
        self.focus_search_entry()
        # Unbind the Enter and Escape keys
        self.master.unbind('<Return>')
        self.master.unbind('<Escape>')

        # Optional: Reset focus to the product list or the edit button
        self.edit_button.focus_set()
    
        # Additionally, you might want to re-bind the Enter key to the edit_button's command
        # if you want to be able to press Enter to switch to edit mode again
        self.master.bind('<Return>', lambda e: self.edit_button.invoke())

    def refresh_and_select_product(self, product_id):
        # Refresh the list of products
        self.combine_and_display_folders()
        
        # Convert the product_id to uppercase for case-insensitive comparison
        product_id_upper = product_id.upper()

        # Find the index of the product that was just edited
        product_index = None
        for index, product_name in enumerate(self.folder_list.get(0, tk.END)):
            # Use .split() to get the first part of the folder name and compare it in uppercase
            if product_name.split()[0].upper() == product_id_upper:
                product_index = index
                break
        
        # If the product is found in the list, select it
        if product_index is not None:
            self.folder_list.selection_set(product_index)
            self.folder_list.see(product_index)  # Ensure the product is visible in the list
            self.folder_list.event_generate("<<ListboxSelect>>")  # Trigger the event to display product details
        self.toggle_edit_mode()
        self.focus_search_entry()

    def get_folder_names_from_db(self):
        self.db_manager.cur.execute("SELECT Folder FROM folder_paths")
        return [row[0] for row in self.db_manager.cur.fetchall()]

    def get_folder_path_from_db(self, product_id):
        # This query assumes that the folder name starts with the product ID followed by a space
        self.db_manager.cur.execute("SELECT Path FROM folder_paths WHERE Folder LIKE ?", (product_id + ' %',))
        result = self.db_manager.cur.fetchone()
        return result[0] if result else None

    def select_excel_database(self):
        filepath = filedialog.askopenfilename(
            title="Select Excel Database",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        if filepath:
            self.excel_manager.filepath = filepath  # Save the filepath to the ExcelManager instance
            xls = pd.ExcelFile(filepath)
            sheet_names = xls.sheet_names
            if sheet_names:
                # Automatically select the first sheet if available
                self.excel_manager.sheet_name = sheet_names[0]  # Save the sheet name to the ExcelManager instance
                self.save_excel_settings(filepath, sheet_names[0])  # Save settings
                self.excel_manager.load_data()  # Load the data
                self.update_excel_label()  # Update the label
        xls = pd.ExcelFile(filepath)
        sheet_names = xls.sheet_names
        self.ask_sheet_name(sheet_names, filepath)  # Pass filepath here

    def update_excel_label(self):
        excel_db_text = f"{self.excel_manager.filepath} - Sheet: {self.excel_manager.sheet_name}"
        self.excel_db_label.config(text=excel_db_text)

    def ask_sheet_name(self, sheet_names, filepath):
        sheet_window = tk.Toplevel(self)
        sheet_window.title("Select a Sheet")

        listbox = tk.Listbox(sheet_window, exportselection=False)
        listbox.pack(padx=10, pady=10)

        # Populate listbox with sheet names
        for sheet in sheet_names:
            listbox.insert('end', sheet)

        # Set the default selection
        default_sheet_index = sheet_names.index(self.default_sheet) if self.default_sheet in sheet_names else 0
        listbox.selection_set(default_sheet_index)
        listbox.activate(default_sheet_index)

        # Bind double-click event to the listbox
        listbox.bind('<Double-1>', lambda event: self.confirm_sheet_selection(event, listbox, filepath))

        confirm_button = tk.Button(sheet_window, text="Confirm", command=lambda: self.confirm_sheet_selection(None, listbox, filepath))
        confirm_button.pack(pady=(0, 10))

        sheet_window.wait_window()

    def confirm_sheet_selection(self, event, listbox, filepath):
        selection_index = listbox.curselection()
        if selection_index:
            selected_sheet = listbox.get(selection_index[0])
            self.select_excel_sheet(selected_sheet, filepath)
            listbox.master.destroy()  # Closes the sheet_window

    def select_excel_sheet(self, selected_sheet, filepath):
        # Code to update the ExcelManager with the new sheet and load data
        self.excel_manager.filepath = filepath
        self.excel_manager.sheet_name = selected_sheet
        self.excel_manager.load_data()
        self.update_excel_label()
        self.save_excel_settings(filepath, selected_sheet)

    def save_excel_settings(self, filepath, sheet_name):
        try:
            with open('excel_and_sheet_path.txt', 'w') as f:
                f.write(f"{filepath}\n{sheet_name}")
            self.update_excel_label()  # Update the label when settings are saved
        except Exception as e:
            messagebox.showerror("Error", f"Unable to save settings: {str(e)}")

    def load_excel_settings(self):
        try:
            with open('excel_and_sheet_path.txt', 'r') as f:
                filepath, sheet_name = f.read().strip().split('\n', 1)
                return filepath, sheet_name
        except FileNotFoundError:
            return None, None
        except Exception as e:
            messagebox.showerror("Error", f"Unable to load settings: {str(e)}")
            return None, None

    def correlate_data(self):
        #print("Correlate button pressed")
        
        filepath, sheet_name = self.load_excel_settings()

        # Check if the Excel settings are properly loaded
        if not filepath or not sheet_name:
            messagebox.showerror("Error", "Excel database settings not found.")
            return

        # Load the data into the ExcelManager instance
        self.excel_manager.filepath = filepath  # Set the filepath
        self.excel_manager.sheet_name = sheet_name  # Set the sheet name
        self.excel_manager.load_data()  # Load the data
        
        try:
            # Load Excel data
            df = pd.read_excel(filepath, sheet_name=sheet_name)
            #print("Excel data loaded successfully.")
            product_ids = df['Product ID'].tolist()
            #print(f"Product IDs from Excel: {product_ids}")
        except Exception as e:
            messagebox.showerror("Error", f"Unable to load Excel file: {str(e)}")
            return
            # Filter out nan values from the product_ids list using pandas notnull function
            
        # Sort the DataFrame based on 'Product ID'
        df_sorted = df.sort_values('Product ID').dropna(subset=['Product ID'])

        # Filter out nan values from the product_ids list
        product_ids = df_sorted['Product ID'].tolist()
        #print(f"Sorted and Filtered Product IDs from Excel: {product_ids}")
        
        missing_docs = []
        for product_id in product_ids:
            folder_path = self.get_folder_path_from_db(str(product_id))
            #print(f"Checking folder for Product ID {product_id}: {folder_path}")
            if folder_path:
                word_docs = [f for f in os.listdir(folder_path) if f.endswith('.docx')]
                #print(f"Word documents in folder: {word_docs}")
                if not word_docs:  # If there's no Word document
                    product_name = df.loc[df['Product ID'] == product_id, 'Product Name'].iloc[0]
                    missing_docs.append((os.path.basename(folder_path), product_id, product_name))


        #print(f"Missing documents: {missing_docs}")
        if missing_docs:
            self.prompt_correlation(missing_docs)
        else:
            messagebox.showinfo("Check complete", "No missing Word documents found.")
        # Filter out nan values from the product_ids list

    def update_links_in_excel(self):
        try:
            with open('excel_and_sheet_path.txt', 'r') as file:
                lines = file.readlines()
                excel_path = lines[0].strip()
                sheet_name = lines[1].strip()

            workbook = openpyxl.load_workbook(excel_path)
            sheet = workbook[sheet_name]

            # Find the index of the columns
            header_row = sheet[1]
            product_name_col_index = None
            order_link_col_index = None

            for index, cell in enumerate(header_row):
                if cell.value == 'Product Name':
                    product_name_col_index = index + 1
                elif cell.value == 'Order Link':
                    order_link_col_index = index + 1

            if product_name_col_index is None or order_link_col_index is None:
                messagebox.showerror("Error", "Necessary columns not found.")
                return

            # Iterate through all the rows and update hyperlinks in 'Order Link' column
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=product_name_col_index):
                product_name_cell = row[product_name_col_index - 1]
                order_link_cell = sheet.cell(row=product_name_cell.row, column=order_link_col_index)
                # Copy only the hyperlink URL
                if product_name_cell.hyperlink:
                    order_link_cell.hyperlink = product_name_cell.hyperlink
                    order_link_cell.value = product_name_cell.hyperlink.target  # Set the cell value to the hyperlink URL

            workbook.save(excel_path)
            messagebox.showinfo("Success", "Links have been updated in the Excel file.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while updating links: {e}")
        self.update_asin_in_excel()

    def update_asin_in_excel(self):
        try:
            with open('excel_and_sheet_path.txt', 'r') as file:
                lines = file.readlines()
                excel_path = lines[0].strip()
                sheet_name = lines[1].strip()

            workbook = openpyxl.load_workbook(excel_path)
            sheet = workbook[sheet_name]

            # Find the index of the columns
            header_row = sheet[1]
            order_link_col_index = None
            asin_col_index = None

            for index, cell in enumerate(header_row):
                if cell.value == 'Order Link':
                    order_link_col_index = index + 1
                elif cell.value == 'ASIN':
                    asin_col_index = index + 1

            if order_link_col_index is None or asin_col_index is None:
                print("Order Link or ASIN columns not found.")  # Debug print
                messagebox.showerror("Error", "Order Link or ASIN columns not found.")
                return

            # Iterate through all the rows and update ASIN based on 'Order Link'
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=order_link_col_index):
                order_link_cell = row[order_link_col_index - 1]
                if order_link_cell.value and '/' in order_link_cell.value:
                    asin_value = order_link_cell.value.split('/')[-1]
                    asin_cell = sheet.cell(row=order_link_cell.row, column=asin_col_index)
                    asin_cell.value = asin_value
                    print(f"Updated ASIN for row {order_link_cell.row}: {asin_value}")  # Debug print

            workbook.save(excel_path)
            print("Excel file saved with updated ASINs.")  # Debug print
            messagebox.showinfo("Success", "ASINs have been updated in the Excel file.")

        except Exception as e:
            print(f"An error occurred while updating ASINs: {e}")  # Debug print
            messagebox.showerror("Error", f"An error occurred while updating ASINs: {e}")
        self.update_to_sell_after_in_excel()

    def update_to_sell_after_in_excel(self):
        try:
            with open('excel_and_sheet_path.txt', 'r') as file:
                lines = file.readlines()
                excel_path = lines[0].strip()
                sheet_name = lines[1].strip()

            workbook = openpyxl.load_workbook(excel_path)
            sheet = workbook[sheet_name]

            # Find the index of the columns
            header_row = sheet[1]
            order_date_col_index = None
            to_sell_after_col_index = None

            for index, cell in enumerate(header_row):
                if cell.value == 'Order Date':
                    order_date_col_index = index + 1
                elif cell.value == 'To Sell After':
                    to_sell_after_col_index = index + 1

            if order_date_col_index is None or to_sell_after_col_index is None:
                print("Order Date or To Sell After columns not found.")  # Debug print
                messagebox.showerror("Error", "Order Date or To Sell After columns not found.")
                return

            # Iterate through all the rows and update 'To Sell After' based on 'Order Date'
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=order_date_col_index):
                order_date_cell = row[order_date_col_index - 1]
                if order_date_cell.value and isinstance(order_date_cell.value, datetime):
                    to_sell_after_date = order_date_cell.value + relativedelta(months=+6)
                    to_sell_after_cell = sheet.cell(row=order_date_cell.row, column=to_sell_after_col_index)
                    to_sell_after_cell.value = to_sell_after_date
                    print(f"Updated To Sell After for row {order_date_cell.row}: {to_sell_after_date}")  # Debug print

            workbook.save(excel_path)
            print("Excel file saved with updated To Sell After dates.")  # Debug print
            messagebox.showinfo("Success", "To Sell After dates have been updated in the Excel file.")

        except Exception as e:
            print(f"An error occurred while updating To Sell After dates: {e}")  # Debug print
            messagebox.showerror("Error", f"An error occurred while updating To Sell After dates: {e}")

    def update_folder_names(self):
        # Load folder paths from folders_paths.txt
        with open("folders_paths.txt", "r") as file:
            lines = file.read().splitlines()
            self.inventory_folder = lines[0]
            self.sold_folder = lines[1]
            self.to_sell_folder = lines[2] if len(lines) > 2 else None
        
        # Load Excel path and sheet name from excel_and_sheet_path.txt
        with open('excel_and_sheet_path.txt', 'r') as f:
            excel_path, sheet_name = f.read().strip().split('\n', 1)
        
        # Load Excel data
        df = pd.read_excel(excel_path, sheet_name)

        print("Starting the folder renaming process...")

        # Iterate over each folder in the inventory, sold, and to sell folders
        for folder_path in [self.inventory_folder, self.sold_folder, self.to_sell_folder]:
            if folder_path and os.path.exists(folder_path):
                # Instead of comparing folder names directly, create a set for more efficient checks

                for item in os.listdir(folder_path):
                    item_path = os.path.join(folder_path, item)
                    if os.path.isdir(item_path):
                        # Extract the presumed product_id from the folder name
                        presumed_product_id = item.split(' ')[0]

                        # Find the matching product_id in the DataFrame
                        product_info = df[df['Product ID'].str.upper() == presumed_product_id.upper()]
                        if not product_info.empty:
                            # Extract the actual product_id and product_name
                            product_id = product_info['Product ID'].iloc[0]
                            product_name = product_info['Product Name'].iloc[0]
                            
                            # Generate the new folder name and sanitize it
                            new_folder_name = self.replace_invalid_chars(f"{product_id} - {product_name}")
                            new_full_path = self.shorten_path(product_id, product_name, folder_path)
                            new_folder_name_from_path = os.path.basename(new_full_path)

                            # Convert both names to a comparable format
                            comparable_item = self.replace_invalid_chars(item).lower().strip()
                            comparable_new_name = new_folder_name_from_path.lower().strip()

                            # Check if the current folder name is already in the correct format
                            if comparable_item == comparable_new_name:
                                print((f"Folder '{item}' already contains the product name.").encode('utf-8', errors='ignore').decode('cp1252', errors='ignore'))
                                continue

                            # Check if the new full path length is within the limit
                            if len(new_full_path) < 260:
                                try:
                                    os.rename(item_path, new_full_path)
                                    print((f"Renamed '{item}' to '{new_folder_name}'").encode('utf-8', errors='ignore').decode('cp1252', errors='ignore'))
                                except OSError as e:
                                    print((f"Error renaming {item_path} to {new_full_path}: {e}").encode('utf-8', errors='ignore').decode('cp1252', errors='ignore'))
                            else:
                                print((f"Skipped renaming {item_path} due to path length restrictions.").encode('utf-8', errors='ignore').decode('cp1252', errors='ignore'))
                        else:
                            print((f"No matching product info found for folder: {item}").encode('utf-8', errors='ignore').decode('cp1252', errors='ignore'))
        self.db_manager.delete_all_folders()
        self.db_manager.setup_database()
        messagebox.showinfo("Done", "Moved and renamed folders")
        
    def replace_invalid_chars(self, filename):
        # Windows filename invalid characters
        invalid_chars = '<>:"/\\|?*'
        for ch in invalid_chars:
            if ch in filename:
                filename = filename.replace(ch, "x")
        return filename

    def shorten_path(self, product_id, product_name, base_path):
        # Windows MAX_PATH is 260 characters
        MAX_PATH = 260
        # Initial maximum length for the product name
        max_name_length = 60

        while max_name_length > 0:
            # Truncate product name to fit
            truncated_product_name = product_name[:max_name_length]

            new_folder_name = f"{product_id} - {truncated_product_name}"
            new_full_path = os.path.join(base_path, new_folder_name)

            # Check if the full path length is within the limit
            if len(new_full_path) <= MAX_PATH:
                return new_full_path
            else:
                # Decrement the maximum name length for the next iteration
                max_name_length -= 1

        # If the loop ends without finding a suitable length, return None or handle appropriately
        print("Unable to shorten the product name sufficiently.")
        return None

    def update_folders_paths(self):
        print("Updating folder paths based on Excel data...")

        # Ensure the Excel file path and sheet name are set
        filepath, sheet_name = self.load_excel_settings()
        if not filepath or not sheet_name:
            print("Excel file path or sheet name is not set.")
            return

        # Load the Excel data
        self.excel_manager.filepath = filepath
        self.excel_manager.sheet_name = sheet_name
        self.excel_manager.load_data()

        # Iterate through Inventory folders
        if self.inventory_folder and os.path.exists(self.inventory_folder):
            for root, dirs, _ in os.walk(self.inventory_folder):
                for dir_name in dirs:
                    product_id = dir_name.split(' ')[0]  # Assuming Product ID is the first part of the name
                    product_info = self.excel_manager.get_product_info(product_id)

                    if product_info:
                        sold_status = product_info.get('Sold')
                        to_sell_after = product_info.get('To Sell After')

                        if sold_status and sold_status.upper() == 'YES':
                            self.move_product_folder(root, dir_name, self.sold_folder)
                        elif sold_status.upper() == 'NO' and self.is_date_today_or_before(to_sell_after):
                            self.move_product_folder(root, dir_name, self.to_sell_folder)
        else:
            print(f"Inventory folder not found: {self.inventory_folder}")
        self.db_manager.delete_all_folders()
        self.db_manager.setup_database()
        self.update_folder_names()

    def move_product_folder(self, current_path, folder_name, target_folder):
        if target_folder and os.path.exists(target_folder):
            full_path = os.path.join(current_path, folder_name)

            # Extracting the part of the folder name before the first '-' and keeping the hyphen
            new_folder_name = folder_name.split('-', 1)[0].strip() + ' -'
            new_full_path = os.path.join(target_folder, new_folder_name)

            try:
                # Check if a folder with the new name already exists in the target directory
                if os.path.exists(new_full_path):
                    print(f"Folder with name '{new_folder_name}' already exists in the target directory.")
                    return

                # Rename and move the folder
                os.rename(full_path, new_full_path)
                print(f"Moved and renamed folder '{folder_name}' to '{new_folder_name}' in '{target_folder}'")
            except Exception as e:
                print(f"Error moving folder '{folder_name}': {e}")
        else:
            print(f"Target folder not found: {target_folder}")

    def is_date_today_or_before(self, date_str):
        if pd.isnull(date_str):
            return False

        # If date_str is a Timestamp, convert it to a datetime object
        if isinstance(date_str, pd.Timestamp):
            to_sell_date = date_str.to_pydatetime().date()
        else:
            try:
                # Assuming date_str is a string in the format "%m/%d/%Y"
                to_sell_date = datetime.strptime(date_str, "%m/%d/%Y").date()
            except ValueError:
                print(f"Invalid date format: {date_str}")
                return False

        return to_sell_date <= datetime.today().date()

    def prompt_correlation(self, missing_docs):
        self.correlate_window = Toplevel(self)
        self.correlate_window.title("Correlate Data")

        self.missing_docs = missing_docs

        # Create a Treeview with columns
        self.correlate_tree = ttk.Treeview(self.correlate_window, columns=('Folder Name', 'Product ID', 'Product Name'), show='headings')
        self.correlate_tree.pack(fill='both', expand=True)

        # Configure the columns
        self.correlate_tree.column('Folder Name', anchor='w', width=150)
        self.correlate_tree.column('Product ID', anchor='center', width=100)
        self.correlate_tree.column('Product Name', anchor='w', width=150)

        # Define the headings
        self.correlate_tree.heading('Folder Name', text='Folder Name', anchor='w')
        self.correlate_tree.heading('Product ID', text='Product ID', anchor='center')
        self.correlate_tree.heading('Product Name', text='Product Name', anchor='w')

        # Add the items to the Treeview
        for i, (folder_name, product_id, product_name) in enumerate(missing_docs):
            self.correlate_tree.insert('', 'end', iid=str(i), values=(folder_name, product_id, product_name))

        # Bind double-click event to an item
        self.correlate_tree.bind('<Double-1>', self.on_item_double_click)
        
        # Adding a Yes to All button
        yes_to_all_button = ttk.Button(self.correlate_window, text="Yes to All", command=self.create_all_word_docs)
        yes_to_all_button.pack()


        exit_button = ttk.Button(self.correlate_window, text="Exit", command=self.exit_correlate_window)
        exit_button.pack()

    def on_item_double_click(self, event):
        item_id = self.correlate_tree.selection()[0]  # Get selected item ID (iid)
        item_values = self.correlate_tree.item(item_id, 'values')
        
        # Convert item values to a doc_data tuple (folder_name, product_id, product_name)
        doc_data = (item_values[0], item_values[1], item_values[2])

        # Call the create_word_doc function with doc_data and the item's iid
        self.create_word_doc(doc_data, item_id)  # show_message is True by default

    def create_all_word_docs(self):
        #print("Create all word docs function called")  # Debug #print statement
        for iid in self.correlate_tree.get_children():
            item_values = self.correlate_tree.item(iid, 'values')
            doc_data = (item_values[0], item_values[1], item_values[2])
            self.create_word_doc(doc_data, iid, show_message=False)
        messagebox.showinfo("Success", "All Word documents have been created.")
        self.correlate_window.destroy()
        self.open_settings_window()

    def create_word_doc(self, doc_data, iid, show_message=True):
        #print("Create word doc function called")  # Debug #print statement
        folder_name, product_id, product_name = doc_data
        folder_path = self.get_folder_path_from_db(str(product_id))

        if folder_path:
            # Here's where you attempt to retrieve the fair market value
            try:
                # Make sure the column names used here match exactly with those in your Excel file
                fair_market_value_series = self.excel_manager.data_frame.loc[self.excel_manager.data_frame['Product ID'] == product_id, 'Fair Market Value']
                if not fair_market_value_series.empty:
                    fair_market_value = fair_market_value_series.iloc[0]
                else:
                    fair_market_value = "N/A"  # Default to "N/A" if the value is not found
                    
                # Retrieving the product link
                order_link_series = self.excel_manager.data_frame.loc[self.excel_manager.data_frame['Product ID'] == product_id, 'Order Link']
                if not order_link_series.empty:
                    order_link = order_link_series.iloc[0]
                else:
                    order_link = "N/A"  # Default to "N/A" if the link is not found
            except Exception as e:
                print(f"Error retrieving data: {e}")  # Debugging #print statement

            doc_path = os.path.join(folder_path, f"{product_id}.docx")
            try:
                doc = Document()
                doc.add_paragraph(f"Product ID: {product_id}")
                doc.add_paragraph(f"Product Name: {product_name}")
                doc.add_paragraph(f"Fair Market Value: ${fair_market_value}")
                doc.add_paragraph(f"Product Link(to get the product description, if needed): {order_link}")
                doc.save(doc_path)
                if show_message:
                    messagebox.showinfo("Document Created", f"Word document for '{product_id}' has been created successfully.")
                # Check if the correlate_tree attribute exists and if there are any items left
                if hasattr(self, 'correlate_tree') and not self.correlate_tree.get_children():
                    self.correlate_window.destroy()
                    self.open_settings_window()
            except Exception as e:
                #print(f"Error creating word doc: {e}")  # Debug #print statement
                messagebox.showerror("Error", f"Failed to create document for Product ID {product_id}: {e}")
        else:
            messagebox.showerror("Error", f"No folder found for Product ID {product_id}")

    def backup_excel_database(self):
        print("Starting the backup process.")
        # Check if the Excel filepath is set
        if not self.excel_manager.filepath:
            print("No Excel filepath is set.")
            return
        
        # Check if inventory folder exists
        if not self.inventory_folder or not os.path.exists(self.inventory_folder):
            print(f"Inventory folder is not set or does not exist: {self.inventory_folder}")
            return
        
        # Define backup folder, which is alongside the inventory folder
        # We obtain the parent directory of the inventory_folder using os.path.dirname
        parent_dir = os.path.dirname(self.inventory_folder)
        backup_folder = os.path.join(parent_dir, "Excel Backups")
        
        try:
            # Create the backup folder if it doesn't exist
            if not os.path.exists(backup_folder):
                os.makedirs(backup_folder)
                print(f"Backup folder '{backup_folder}' created.")
            else:
                print(f"Backup folder '{backup_folder}' already exists.")
            
            # Generate backup file name
            date_time_str = datetime.now().strftime("%Y-%m-%d - %H-%M-%S")
            backup_filename = f"Backup of {date_time_str}.xlsx"
            backup_path = os.path.join(backup_folder, backup_filename)
            
            # Copy the Excel file to the backup location
            shutil.copy2(self.excel_manager.filepath, backup_path)
            print(f"Backup created at: {backup_path}")
            
            # Double check if the file was actually created
            if not os.path.isfile(backup_path):
                raise FileNotFoundError(f"Backup file not found after copy operation: {backup_path}")
        except Exception as e:
            print(f"Failed to create backup: {e}")
            raise  # Reraise the exception to see the full traceback

    def __del__(self):
        self.db_manager.conn.close()

def on_close(app, root):
    print("Closing the application and attempting to backup the database.")
    if hasattr(app, 'excel_manager') and app.excel_manager.filepath:
        print(f"Excel file path at time of backup: {app.excel_manager.filepath}")
        try:
            app.backup_excel_database()  # Perform the backup
            print("Backup should now be complete.")
        except Exception as e:
            print(f"An error occurred during backup: {e}")
    else:
        print("Excel manager not set or no filepath available.")
    root.destroy()  # Call the destroy method to close the application

def exit_application(app, root):
    on_close(app)  # Perform backup
    root.destroy()  # Exit the application

def main():
    root = tk.Tk()
    root.title("Improved Inventory Manager")
    app = Application(master=root)
    root.state('zoomed')
    app.excel_manager.filepath, _ = app.load_excel_settings()


    # Use a lambda to pass 'app' and 'root' to the 'on_close' function
    root.protocol("WM_DELETE_WINDOW", lambda: on_close(app, root))
    
    app.mainloop()

if __name__ == '__main__':
    main()