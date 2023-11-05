
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import pandas as pd
from docx import Document
import sqlite3
from tkinter import END
from tkinter import Toplevel

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

    def get_product_info(self, product_id):
        if self.data_frame is not None:
            # Convert both the product_id and the 'Product ID' column to lower case for comparison
            query_result = self.data_frame[self.data_frame['Product ID'].str.lower() == product_id.lower()]
            if not query_result.empty:
                return query_result.iloc[0].to_dict()
        return None

class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.db_manager = DatabaseManager()
        self.excel_manager = ExcelManager()
        self.edit_mode = False  # Add this line to initialize the edit_mode attribute
        self.folder_to_scan = None
        self.sold_folder = None
        self.products_to_sell_folder = None
        self.pack(fill='both', expand=True)
        
        # Make sure you call this before combining and displaying folders
        self.Main_Window_Widgets() 
        
        # Now it's safe to load settings and combine folders since the list widget is created
        self.load_settings()
        self.combine_and_display_folders()

    def load_settings(self):
        # Load settings
        try:
            with open("folders_settings.txt", "r") as file:
                lines = file.read().splitlines()
                self.folder_to_scan = lines[0]
                self.sold_folder = lines[1]
                self.products_to_sell_folder = lines[2] if len(lines) > 2 else None
                # ... The rest of your settings loading code ...
        except FileNotFoundError:
            pass
        # Here you could handle the situation if the file is not found, like setting default paths or prompting the user.

    def save_settings(self):
        # This function is called after selecting the source and sold folders
        # Update the table with the new paths
        self.db_manager.cur.execute('''
            UPDATE folder_paths SET Path = ? WHERE Folder = 'Root Folder'
        ''', (self.folder_to_scan,))
        self.db_manager.cur.execute('''
            UPDATE folder_paths SET Path = ? WHERE Folder = 'Sold'
        ''', (self.sold_folder,))
        self.db_manager.conn.commit()

    def check_and_update_product_list(self):
        if not self.search_entry.get():  # Check if the search entry is empty
            folder_count = len(next(os.walk(self.folder_to_scan))[1])  # Count folders in the directory
            list_count = self.folder_list.size()  # Count items in the Listbox

            if folder_count != list_count:
                self.display_folders(self.folder_to_scan)  # Update the list items with folder names

            # Schedule this method to be called again after 10000 milliseconds (10 seconds)
            self.after(10000, self.check_and_update_product_list)

    def Main_Window_Widgets(self):
        self.top_frame = tk.Frame(self)
        self.top_frame.pack(fill='x')

        self.settings_button = tk.Button(self.top_frame, text='Settings', command=self.open_settings)
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
        self.product_frame.pack(side='right', fill='both', expand=True)

        self.sold_var = tk.BooleanVar()
        self.sold_checkbutton = tk.Checkbutton(self.product_frame, text='Sold', variable=self.sold_var)
        
        self.save_button = tk.Button(self.product_frame, text='Save', command=self.save, state='disabled')
        self.save_button.grid(row=0, column=8, sticky='w', padx=200, pady=0)

        # Row 0
        self.edit_button = tk.Button(self.product_frame, text="Edit", command=self.toggle_edit_mode)
        self.edit_button.grid(row=0, column=8, sticky='w', padx=235, pady=0)
        
        self.cancelled_order_var = tk.BooleanVar()
        self.cancelled_order_checkbutton = tk.Checkbutton(self.product_frame, text='Cancelled Order', variable=self.cancelled_order_var, state='disabled')
        self.cancelled_order_checkbutton.grid(row=0, column=6, sticky='w', padx=200, pady=0)
        
        self.order_date_var = tk.StringVar()
        self.order_date_label = tk.Label(self.product_frame, text='Order Date')
        self.order_date_label.grid(row=10, column=0, sticky='w', padx=0, pady=0)
        self.order_date_var.trace("w", self.update_to_sell_after)

        # Row 1
        self.damaged_var = tk.BooleanVar()
        self.damaged_checkbutton = tk.Checkbutton(self.product_frame, text='Damaged', variable=self.damaged_var, state='disabled')
        self.damaged_checkbutton.grid(row=1, column=6, sticky='w', padx=200, pady=0)
        
        self.order_date_entry = DateEntry(self.product_frame, textvariable=self.order_date_var, state='disabled')
        self.order_date_entry.grid(row=11, column=0, sticky='w', padx=0, pady=0)

        # Row 2
        self.personal_var = tk.BooleanVar()
        self.personal_checkbutton = tk.Checkbutton(self.product_frame, text='Personal', variable=self.personal_var, state='disabled')
        self.personal_checkbutton.grid(row=2, column=6, sticky='w', padx=200, pady=0)
        
        self.to_sell_after_var = tk.StringVar()
        self.to_sell_after_label = tk.Label(self.product_frame, text='To Sell After')
        self.to_sell_after_label.grid(row=12, column=0, sticky='w', padx=0, pady=0)

        # Row 3
        self.reviewed_var = tk.BooleanVar()
        self.reviewed_checkbutton = tk.Checkbutton(self.product_frame, text='Reviewed', variable=self.reviewed_var, state='disabled')
        self.reviewed_checkbutton.grid(row=3, column=6, sticky='w', padx=200, pady=0)
        
        self.to_sell_after_entry = DateEntry(self.product_frame, textvariable=self.to_sell_after_var, state='disabled')
        self.to_sell_after_entry.grid(row=13, column=0, sticky='w', padx=0, pady=0)

        # Row 4
        self.pictures_downloaded_var = tk.BooleanVar()
        self.pictures_downloaded_checkbutton = tk.Checkbutton(self.product_frame, text='Pictures Downloaded', variable=self.pictures_downloaded_var, state='disabled')
        self.pictures_downloaded_checkbutton.grid(row=4, column=6, sticky='w', padx=200, pady=0)
        self.payment_type_var = tk.StringVar()
        self.payment_type_label = tk.Label(self.product_frame, text='Payment Type')
        self.payment_type_label.grid(row=4, column=0, sticky='w', padx=0, pady=0)

        # Row 5
        self.uploaded_to_site_var = tk.BooleanVar()
        self.uploaded_to_site_checkbutton = tk.Checkbutton(self.product_frame, text='Uploaded to Site', variable=self.uploaded_to_site_var, state='disabled')
        self.uploaded_to_site_checkbutton.grid(row=5, column=6, sticky='w', padx=200, pady=0)
        self.payment_type_combobox = ttk.Combobox(self.product_frame, textvariable=self.payment_type_var, state='disabled')
        self.payment_type_combobox['values'] = ('Cash', 'ATH Movil')  # default options
        self.payment_type_combobox.grid(row=5, column=0, sticky='w', padx=0, pady=0)


        self.product_folder_var = tk.StringVar()
        self.product_folder_label = tk.Label(self.product_frame, text='Product Folder')
        self.product_folder_label.grid(row=6, column=0, sticky='w', padx=0, pady=0)
        
        self.product_folder_entry = tk.Entry(self.product_frame, textvariable=self.product_folder_var, state='disabled')
        self.product_folder_entry.grid(row=7, column=0, sticky='w', padx=0, pady=0)
        
        self.sold_var.set(False)
        self.sold_checkbutton = tk.Checkbutton(self.product_frame, text='Sold', variable=self.sold_var, state='disabled')
        self.sold_checkbutton.grid(row=6, column=6, sticky='w', padx=200, pady=0)

        self.asin_var = tk.StringVar()
        self.asin_label = tk.Label(self.product_frame, text='ASIN')
        self.asin_label.grid(row=8, column=0, sticky='w', padx=0, pady=0)
        
        self.asin_entry = tk.Entry(self.product_frame, textvariable=self.asin_var, state='disabled')
        self.asin_entry.grid(row=9, column=0, sticky='w', padx=0, pady=0)

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

        # Note: Product Image requires a different approach

        self.fair_market_value_var = tk.StringVar()
        self.fair_market_value_label = tk.Label(self.product_frame, text='Fair Market Value')
        self.fair_market_value_label.grid(row=16, column=0, sticky='w', padx=0, pady=0)
        
        self.fair_market_value_entry = tk.Entry(self.product_frame, textvariable=self.fair_market_value_var, state='disabled')
        self.fair_market_value_entry.grid(row=17, column=0, sticky='w', padx=0, pady=0)

        self.order_details_var = tk.StringVar()
        self.order_details_label = tk.Label(self.product_frame, text='Order Details')
        self.order_details_label.grid(row=18, column=0, sticky='w', padx=0, pady=0)
        
        self.order_details_entry = tk.Entry(self.product_frame, textvariable=self.order_details_var, state='disabled')
        self.order_details_entry.grid(row=19, column=0, sticky='w', padx=0, pady=0)

        self.order_link_var = tk.StringVar()
        self.order_link_label = tk.Label(self.product_frame, text='Order Link')
        self.order_link_label.grid(row=20, column=0, sticky='w', padx=0, pady=0)
        
        self.order_link_entry = tk.Entry(self.product_frame, textvariable=self.order_link_var, state='disabled')
        self.order_link_entry.grid(row=21, column=0, sticky='w', padx=0, pady=0)

        self.sold_price_var = tk.StringVar()
        self.sold_price_label = tk.Label(self.product_frame, text='Sold Price')
        self.sold_price_label.grid(row=22, column=0, sticky='w', padx=0, pady=0)
        
        self.sold_price_entry = tk.Entry(self.product_frame, textvariable=self.sold_price_var, state='disabled')
        self.sold_price_entry.grid(row=23, column=0, sticky='w', padx=0, pady=0)
        
        # Load settings
        try:
            with open("folders_settings.txt", "r") as file:
                lines = file.read().splitlines()
                self.folder_to_scan = lines[0]
                self.sold_folder = lines[1]
                self.products_to_sell_folder = lines[2] if len(lines) > 2 else None
                if self.folder_to_scan:  # Check if folder_to_scan is defined
                    self.display_folders(self.folder_to_scan)
        except FileNotFoundError:
            pass


        self.search_entry.focus_set()

    def focus_search_entry(self):
        self.search_entry.focus_set()

    def open_settings(self):
        if hasattr(self, 'settings_window') and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return
        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("Settings")
        self.settings_window.update()  # This updates the window and calculates sizes
        window_width = self.settings_window.winfo_width()  # Gets the width of the window
        
        # Configure the grid layout to not expand the button rows
        self.settings_window.grid_rowconfigure(0, weight=0)
        self.settings_window.grid_rowconfigure(1, weight=0)
        self.settings_window.grid_rowconfigure(2, weight=0)  # Add this line for your new button row
        self.settings_window.grid_rowconfigure(3, weight=1)  # This row will expand and push content to the top
        
        # Configure the grid columns (if necessary)
        self.settings_window.grid_columnconfigure(0, weight=1)

        self.folder_to_scan_frame = tk.Frame(self.settings_window)
        self.folder_to_scan_frame.grid(row=0, column=0, sticky='w')
        self.folder_to_scan_button = tk.Button(self.folder_to_scan_frame, text="Choose Root Inventory Folder", command=self.choose_folder_to_scan)
        self.folder_to_scan_button.grid(row=1, column=0, padx=(window_width//4, 0))  # Half the remaining space to the left
        self.folder_to_scan_label = tk.Label(self.folder_to_scan_frame, text=self.folder_to_scan if hasattr(self, 'folder_to_scan') else "Not chosen")
        self.folder_to_scan_label.grid(row=1, column=1, padx=(0, window_width//4), sticky='ew')  # Half the remaining space to the right

        self.sold_folder_frame = tk.Frame(self.settings_window)
        self.sold_folder_frame.grid(row=1, column=0, sticky='w')
        self.sold_folder_button = tk.Button(self.sold_folder_frame, text="Choose Sold Inventory Folder", command=self.choose_sold_folder)
        self.sold_folder_button.grid(row=2, column=0, padx=(window_width//4, 0))
        self.sold_folder_label = tk.Label(self.sold_folder_frame, text=self.sold_folder if hasattr(self, 'sold_folder') else "Not chosen")
        self.sold_folder_label.grid(row=2, column=1, padx=(0, window_width//4), sticky='ew')

        # Inside the open_settings method of Application class after existing setup code for other buttons

        # Choose Folder with Products to Sell Button and Label
        self.products_to_sell_folder_frame = tk.Frame(self.settings_window)
        self.products_to_sell_folder_frame.grid(row=3, column=0, sticky='w')  # Adjust the row index as needed
        self.products_to_sell_folder_button = tk.Button(
            self.products_to_sell_folder_frame,
            text="Choose Folder with Products to Sell",
            command=self.choose_products_to_sell_folder
        )
        self.products_to_sell_folder_button.grid(row=3, column=0, padx=(window_width // 4, 0))
        self.products_to_sell_folder_label = tk.Label(
            self.products_to_sell_folder_frame,
            text=self.products_to_sell_folder if self.products_to_sell_folder else "Not chosen"
        )
        self.products_to_sell_folder_label.grid(row=3, column=1, padx=(0, window_width // 4), sticky='ew')

        # Add a new frame for the Excel database selection
        # Load settings
        self.default_filepath, self.default_sheet = self.load_excel_settings()

        # Excel Database Selection frame and button
        self.excel_db_frame = tk.Frame(self.settings_window)
        self.excel_db_frame.grid(row=2, column=0, sticky='w')  # Adjust row as needed
        excel_db_text = f"{self.default_filepath} - Sheet: {self.default_sheet}" if self.default_filepath and self.default_sheet else "Not chosen"
        self.excel_db_label = tk.Label(self.excel_db_frame, text=excel_db_text)
        self.excel_db_button = tk.Button(self.excel_db_frame, text="Select Excel Database", command=self.select_excel_database)
        self.excel_db_button.grid(row=3, column=0, padx=(window_width//4, 0))
        self.excel_db_label.grid(row=3, column=1, padx=(0, window_width//4), sticky='ew')
        
        # Add a new button for "Correlate new data" functionality
        self.correlate_button = tk.Button(self.settings_window, text="Create Word Files for Products", command=self.correlate_data)
        # Adjust the row index accordingly to place the new button
        self.correlate_button.grid(row=4, column=0, padx=(window_width//4, 0), sticky='w')


        self.back_button = tk.Button(self.settings_window, text="<- Back", command=self.back_to_main)
        self.back_button.grid(row=0, column=0, sticky='nw')  # Change this line to place the back button in the fourth row
        
        self.combine_and_display_folders()

        self.master.withdraw()

    def exit_correlate_window(self):
        self.correlate_window.destroy()
        self.open_settings()

    def back_to_main(self):
        self.settings_window.destroy()
        self.master.deiconify()
        self.master.state('zoomed')
        
        # Load settings again in case they were changed
        self.load_settings()
        
        # Refresh the folder list with the updated settings
        self.combine_and_display_folders()
        
        self.focus_search_entry()


    def choose_folder_to_scan(self):
        folder_to_scan = filedialog.askdirectory()
        if folder_to_scan:
            self.folder_to_scan = folder_to_scan
            self.folder_to_scan_label.config(text=folder_to_scan)  # Update the label directly
            self.save_settings()
            self.display_folders(self.folder_to_scan)

    def display_folders(self, folder_to_scan):
        self.folder_list.delete(0, END)
        self.db_manager.cur.execute("DELETE FROM folder_paths")  # Use the cursor from db_manager
        for root, dirs, files in os.walk(folder_to_scan):
            if not dirs:
                name = os.path.basename(root)
                path = root
                self.db_manager.cur.execute("INSERT OR REPLACE INTO folder_paths VALUES (?, ?)", (name, path))
        self.db_manager.conn.commit()
        for folder in sorted(self.get_folder_names_from_db()):
            self.folder_list.insert(END, folder)

    def combine_and_display_folders(self):
        # Clear the folder list first
        self.folder_list.delete(0, tk.END)
        
        # Combine the folders from all three paths
        combined_folders = []
        for folder_path in [self.folder_to_scan, self.sold_folder, self.products_to_sell_folder]:
            if folder_path and os.path.exists(folder_path):
                combined_folders.extend(next(os.walk(folder_path))[1])  # Get the directory names
        
        # Deduplicate folder names
        unique_folders = list(set(combined_folders))
        
        # Insert the unique folders into the list
        for folder in sorted(unique_folders):
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

    def choose_products_to_sell_folder(self):
        self.products_to_sell_folder = filedialog.askdirectory()
        if self.products_to_sell_folder:
            self.products_to_sell_folder_label.config(text=self.products_to_sell_folder)
            self.save_settings()  # Save the settings including the new folder path


    def save_settings(self):
        # Here you will gather all the paths and write them to the settings.txt file
        with open("folders_settings.txt", "w") as file:
            file.write(f"{self.folder_to_scan}\n{self.sold_folder}\n{self.products_to_sell_folder}")


    def search(self, event):
        search_terms = self.search_entry.get().split()  # Split the search string into words
        if search_terms:
            self.folder_list.delete(0, END)  # Clear the current list

            # Walk the directory tree from the folder_to_scan path
            for root, dirs, files in os.walk(self.folder_to_scan):
                # Check if 'dirs' is empty, meaning 'root' is a leaf directory
                if not dirs:
                    folder_name = os.path.basename(root)  # Get the name of the leaf directory
                    # Check if all search terms are in the folder name (case insensitive)
                    if all(term.lower() in folder_name.lower() for term in search_terms):
                        self.folder_list.insert(END, folder_name)
        else:
            self.display_folders(self.folder_to_scan)  # If the search box is empty, display all folders

    def display_product_details(self, event):
        if self.edit_mode:
            if messagebox.askyesno("Unsaved changes", "You have unsaved changes. Do you want to save them?"):
                self.save()
            else:
                self.toggle_edit_mode()  # Reset edit mode
                
        # Get the index of the selected item
        selection = self.folder_list.curselection()
        if not selection:
            return  # No item selected
        index = selection[0]
        selected_folder_name = self.folder_list.get(index)
        selected_product_id = selected_folder_name.split(' ')[0].lower()  # Assuming the product ID is at the beginning

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
                    self.uploaded_to_site_var.set(self.excel_value_to_bool(product_info.get('Uploaded to Site')))
                    self.sold_var.set(self.excel_value_to_bool(product_info.get('Sold')))
                    
                    # For each field, check if the value is NaN using pd.isnull and set it to an empty string if it is
                    self.asin_var.set('' if pd.isnull(product_info.get('ASIN')) else product_info.get('ASIN', ''))
                    self.product_id_var.set('' if pd.isnull(product_info.get('Product ID')) else product_info.get('Product ID', ''))
                    self.product_folder_var.set('' if pd.isnull(product_info.get('Product Folder')) else product_info.get('Product Folder', ''))
                    self.to_sell_after_var.set('' if pd.isnull(product_info.get('To Sell After')) else product_info.get('To Sell After', ''))
                    # ... handle the product image ...
                    self.product_name_var.set('' if pd.isnull(product_info.get('Product Name')) else product_info.get('Product Name', ''))
                    
                    # Get the order date and format it
                    order_date = product_info.get('Order Date', '')
                    if isinstance(order_date, datetime):
                        formatted_order_date = order_date.strftime('%m/%d/%Y')
                    elif isinstance(order_date, str):
                        formatted_order_date = order_date  # Assuming the string is already in the correct format
                    else:
                        formatted_order_date = ''  # If it's neither a datetime object nor a string, set it to an empty string

                    self.order_date_var.set(formatted_order_date)
                    self.fair_market_value_var.set('' if pd.isnull(product_info.get('Fair Market Value')) else product_info.get('Fair Market Value', ''))
                    self.order_details_var.set('' if pd.isnull(product_info.get('Order Details')) else product_info.get('Order Details', ''))
                    self.order_link_var.set('' if pd.isnull(product_info.get('Order Link')) else product_info.get('Order Link', ''))
                    self.sold_price_var.set('' if pd.isnull(product_info.get('Sold Price')) else product_info.get('Sold Price', ''))
                    self.payment_type_var.set('' if pd.isnull(product_info.get('Payment Type')) else product_info.get('Payment Type', ''))
                    # ... continue with other fields as needed ...
                    # Add code here to populate the Sold Date and other date-related fields, if applicable
                else:
                    self.cancelled_order_var.set(False)
                    self.damaged_var.set(False)
                    self.personal_var.set(False)
                    self.reviewed_var.set(False)
                    self.pictures_downloaded_var.set(False)
                    self.uploaded_to_site_var.set(False)
                    self.sold_var.set(False)
                    
                    # Populate the widgets with the matched data
                    self.asin_var.set('')
                    self.product_id_var.set('')
                    self.product_folder_var.set('')
                    self.to_sell_after_var.set('')
                    # Add code here to handle the product image, if applicable
                    self.product_name_var.set('Product not found in Excel.')
                    self.order_date_var.set('')
                    self.fair_market_value_var.set('')
                    self.order_details_var.set('')
                    self.order_link_var.set('')
                    self.sold_price_var.set('')
                    self.payment_type_var.set('')

            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")
                print(f"Error retrieving product details: {e}")
        else:
            messagebox.showerror("Error", "Excel file path or sheet name is not set.")
        
        # Any other code you want to execute when displaying product details, such as configuring widget states

    def excel_value_to_bool(self, value):
        # Check for NaN explicitly and return False if found
        if pd.isnull(value):
            return False
        if isinstance(value, str):
            return value.strip().lower() in ['yes', 'true', '1']
        elif isinstance(value, (int, float)):
            return bool(value)
        return False

    def update_to_sell_after(self, *args):
        order_date_str = self.order_date_var.get()
        if order_date_str:
            try:
                # If the date is in the format 'mm/dd/yy', such as '2/15/23'
                if len(order_date_str.split('/')[-1]) == 2:  # If the year is two digits
                    order_date = datetime.strptime(order_date_str, "%m/%d/%y")
                else:  # If the year is four digits
                    order_date = datetime.strptime(order_date_str, "%m/%d/%Y")

                to_sell_after = order_date + relativedelta(months=6)
                self.to_sell_after_var.set(to_sell_after.strftime("%m/%d/%Y"))
            except ValueError as e:
                messagebox.showerror("Error", f"Incorrect date format: {e}")

    def toggle_edit_mode(self):
        # Toggle the edit mode
        self.edit_mode = not self.edit_mode
        
        # Set the state based on the new edit mode
        state = 'normal' if self.edit_mode else 'disabled'
        self.sold_checkbutton.config(state=state)
        self.cancelled_order_checkbutton.config(state=state)
        self.damaged_checkbutton.config(state=state)
        self.personal_checkbutton.config(state=state)
        self.reviewed_checkbutton.config(state=state)
        self.pictures_downloaded_checkbutton.config(state=state)
        self.uploaded_to_site_checkbutton.config(state=state)
        self.order_date_entry.config(state=state)
        self.to_sell_after_entry.config(state=state)
        self.payment_type_combobox.config(state=state)
        self.asin_entry.config(state=state)
        self.product_id_entry.config(state=state)
        self.product_name_entry.config(state=state)
        self.product_folder_entry.config(state=state)
        self.fair_market_value_entry.config(state=state)
        self.order_details_entry.config(state=state)
        self.order_link_entry.config(state=state)
        self.sold_price_entry.config(state=state)
        self.save_button.config(state=state)


    def save(self):
        # Get the product ID from the input, make sure it's a string, and strip whitespace.
        product_id = self.product_id_var.get().strip()
        
        # Convert the Product ID to lowercase for a case-insensitive comparison, if necessary.
        product_id = product_id.lower()
        
        # Ensure that the Excel file path and sheet name are set.
        filepath, sheet_name = self.load_excel_settings()
        
        if not filepath or not sheet_name:
            messagebox.showerror("Error", "Excel file path or sheet name is not set.")
            return
        
        # Load the Excel data into the DataFrame.
        self.excel_manager.load_data()
        
        # Find the row in the DataFrame that matches the Product ID.
        # This is assuming that the Product ID column in Excel is formatted similarly to the input.
        matching_row = self.excel_manager.data_frame[self.excel_manager.data_frame['Product ID'].str.lower() == product_id]
        
        if matching_row.empty:
            messagebox.showinfo("Info", f"Product ID '{product_id}' not found in the Excel database.")
            return
        
        # Assuming you found the matching row, get the index.
        row_index = matching_row.index[0]
        
        # Collect the data from the form.
        product_data = {
            'Cancelled Order': 'YES' if self.cancelled_order_var.get() else 'NO',
            'Damaged': 'YES' if self.damaged_var.get() else 'NO',
            'Personal': 'YES' if self.personal_var.get() else 'NO',
            'Reviewed': 'YES' if self.reviewed_var.get() else 'NO',
            'Pictures Downloaded': 'YES' if self.pictures_downloaded_var.get() else 'NO',
            'Uploaded to Site': 'YES' if self.uploaded_to_site_var.get() else 'NO',
            'Sold': 'YES' if self.sold_var.get() else 'NO',
            'ASIN': self.asin_var.get(),
            'Product Folder': self.product_folder_var.get(),
            'To Sell After': self.to_sell_after_var.get(),
            'Product Name': self.product_name_var.get(),
            'Fair Market Value': self.fair_market_value_var.get(),
            'Order Details': self.order_details_var.get(),
            'Order Link': self.order_link_var.get(),
            'Sold Price': self.sold_price_var.get(),
            'Payment Type': self.payment_type_var.get(),
            # ... and so on for the rest of your form fields.
        }
        
        # Update the DataFrame with the data collected from the form.
        for key, value in product_data.items():
            self.excel_manager.data_frame.at[row_index, key] = value
        
        # Save the updated DataFrame back to the Excel file.
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                self.excel_manager.data_frame.to_excel(writer, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Success", "Product information updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save changes to Excel file: {e}")
        
        # Reset the form and any state as necessary.
        self.edit_mode = False
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

    def __del__(self):
        self.db_manager.conn.close()

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
            with open('excel_db_settings.txt', 'w') as f:
                f.write(f"{filepath}\n{sheet_name}")
            self.update_excel_label()  # Update the label when settings are saved
        except Exception as e:
            messagebox.showerror("Error", f"Unable to save settings: {str(e)}")

    def load_excel_settings(self):
        try:
            with open('excel_db_settings.txt', 'r') as f:
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
        self.open_settings()

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
            except Exception as e:
                print(f"Error retrieving fair market value: {e}")  # Debugging print statement
                fair_market_value = "N/A"

            doc_path = os.path.join(folder_path, f"{product_id}.docx")
            try:
                doc = Document()
                doc.add_paragraph(f"Product ID: {product_id}")
                doc.add_paragraph(f"Product Name: {product_name}")
                doc.add_paragraph(f"Fair Market Value: {fair_market_value}")
                doc.save(doc_path)
                if show_message:
                    messagebox.showinfo("Document Created", f"Word document for '{product_id}' has been created successfully.")
                self.correlate_tree.delete(iid)
            except Exception as e:
                #print(f"Error creating word doc: {e}")  # Debug #print statement
                messagebox.showerror("Error", f"Failed to create document for Product ID {product_id}: {e}")
        else:
            messagebox.showerror("Error", f"No folder found for Product ID {product_id}")
        
        # After creating the document, check if there are any items left
        if not self.correlate_tree.get_children():
            # If the Treeview is empty, close the 'Correlate Data' window and open the 'Settings' window
            self.correlate_window.destroy()
            self.open_settings()

def main():

    root = tk.Tk()
    root.title("Improved Inventory Manager")
    app = Application(master=root)
    app.mainloop()

if __name__ == '__main__':
    main()
