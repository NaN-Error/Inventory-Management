import os
import shutil
import tkinter as tk
import sqlite3
from tkinter import messagebox, filedialog, LEFT, Y, BOTH, END
from datetime import timedelta
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime
from dateutil.relativedelta import relativedelta

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.after(10000, self.check_and_update_product_list)  # Start the periodic check
        self.initialize_database()
        
        self.pack(fill='both', expand=True)
        self.create_widgets()
        self.edit_mode = False
        self.focus_search_entry()
        
        
    def initialize_database(self):
        # Connect to the SQLite database
        self.conn = sqlite3.connect('inventory_management.db')
        self.cur = self.conn.cursor()

        # Create a new table with the specified columns
        self.cur.execute('''
            CREATE TABLE IF NOT EXISTS folder_paths (
                Folder TEXT PRIMARY KEY,
                Path TEXT
            )
        ''')
        self.conn.commit()

    def save_settings(self):
        # This function is called after selecting the source and sold folders
        # Update the table with the new paths
        self.cur.execute('''
            UPDATE folder_paths SET Path = ? WHERE Folder = 'Root Folder'
        ''', (self.folder_to_scan,))
        self.cur.execute('''
            UPDATE folder_paths SET Path = ? WHERE Folder = 'Sold'
        ''', (self.sold_folder,))
        self.conn.commit()

        
    def check_and_update_product_list(self):
        if not self.search_entry.get():  # Check if the search entry is empty
            folder_count = len(next(os.walk(self.folder_to_scan))[1])  # Count folders in the directory
            list_count = self.folder_list.size()  # Count items in the Listbox

            if folder_count != list_count:
                self.display_folders(self.folder_to_scan)  # Update the list items with folder names

            # Schedule this method to be called again after 10000 milliseconds (10 seconds)
            self.after(10000, self.check_and_update_product_list)

    def create_widgets(self):
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

        self.product_frame = tk.Frame(self.bottom_frame, bg='gray')
        self.product_frame.pack(side='right', fill='both', expand=True)

        self.sold_var = tk.BooleanVar()
        self.sold_checkbutton = tk.Checkbutton(self.product_frame, text='Sold', variable=self.sold_var)
        self.save_button = tk.Button(self.product_frame, text='Save', command=self.save)
        
        # Load settings
        try:
            with open("settings.txt", "r") as file:
                self.folder_to_scan, self.sold_folder = file.read().splitlines()
                if hasattr(self, 'folder_to_scan'):  # Check if folder_to_scan is defined
                    self.display_folders(self.folder_to_scan)
        except FileNotFoundError:
            pass

        self.search_entry.focus_set()

    def focus_search_entry(self):
        self.search_entry.focus_set()

    def open_settings(self):
        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("Settings")
        # Size the window to 50% of the screen size
        window_width = int(self.settings_window.winfo_screenwidth() * 0.5)
        window_height = int(self.settings_window.winfo_screenheight() * 0.5)
        self.settings_window.geometry(f"{window_width}x{window_height}")

        self.settings_window.columnconfigure(0, weight=1)
        self.settings_window.columnconfigure(1, weight=3)
        self.settings_window.rowconfigure(0, weight=0)  # Change this line to not make the first row expand
        self.settings_window.rowconfigure(1, weight=0)  # Change this line to not make the second row expand
        self.settings_window.rowconfigure(2, weight=1)  # This row will expand and push the back button to the bottom
        self.settings_window.state('zoomed')

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

        self.back_button = tk.Button(self.settings_window, text="<- Back", command=self.back_to_main)
        self.back_button.grid(row=0, column=0, sticky='nw')  # Change this line to place the back button in the fourth row

        self.master.withdraw()

    def back_to_main(self):
        self.settings_window.destroy()
        self.master.deiconify()
        self.master.state('zoomed')  # Add this line
        if hasattr(self, 'folder_to_scan'):  # Check if folder_to_scan is defined
            self.display_folders(self.folder_to_scan)
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
        self.cur.execute("DELETE FROM folder_paths")  # Clear the previous folder paths in the database
        for root, dirs, files in os.walk(folder_to_scan):
            if not dirs:
                name = os.path.basename(root)
                path = root
                self.cur.execute("INSERT OR REPLACE INTO folder_paths VALUES (?, ?)", (name, path))
        self.conn.commit()
        for folder in sorted(self.get_folder_names_from_db()):
            self.folder_list.insert(END, folder)

    def choose_sold_folder(self):
        self.sold_folder = filedialog.askdirectory()
        if self.sold_folder:
            self.sold_folder_label.config(text=self.sold_folder)  # Update the label directly
            self.save_settings()


        # Update the Sold Folder path
        self.cur.execute('''
            INSERT INTO folder_paths (Folder, Path) VALUES ('Sold', ?)
            ON CONFLICT(Folder) DO UPDATE SET Path = excluded.Path;
        ''', (self.sold_folder,))

        self.conn.commit()
    def save_settings(self):
        if getattr(self, 'folder_to_scan', None) is not None and getattr(self, 'sold_folder', None) is not None:
            with open("settings.txt", "w") as file:
                file.write(f"{self.folder_to_scan}\n{self.sold_folder}")

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

        self.save_button.grid(row=0, column=8, sticky='w', padx=200, pady=0)
        self.edit_mode = False
        selected_product = self.folder_list.get(self.folder_list.curselection())

        # Row 0
        self.edit_button = tk.Button(self.product_frame, text="Edit", command=self.toggle_edit_mode)
        self.edit_button.grid(row=0, column=8, sticky='w', padx=235, pady=0)
        
        self.cancelled_order_var = tk.BooleanVar()
        self.cancelled_order_checkbutton = tk.Checkbutton(self.product_frame, text='Cancelled Order', variable=self.cancelled_order_var, state='disabled')
        self.cancelled_order_checkbutton.grid(row=0, column=6, sticky='w', padx=200, pady=0)
        
        self.order_date_var = tk.StringVar()
        self.order_date_label = tk.Label(self.product_frame, text='Order Date')
        self.order_date_label.grid(row=0, column=0, sticky='w', padx=0, pady=0)
        self.order_date_var.trace("w", self.update_to_sell_after)


        # Row 1
        self.damaged_var = tk.BooleanVar()
        self.damaged_checkbutton = tk.Checkbutton(self.product_frame, text='Damaged', variable=self.damaged_var, state='disabled')
        self.damaged_checkbutton.grid(row=1, column=6, sticky='w', padx=200, pady=0)
        self.order_date_entry = DateEntry(self.product_frame, textvariable=self.order_date_var, state='disabled')
        self.order_date_entry.grid(row=1, column=0, sticky='w', padx=0, pady=0)

        # Row 2
        self.personal_var = tk.BooleanVar()
        self.personal_checkbutton = tk.Checkbutton(self.product_frame, text='Personal', variable=self.personal_var, state='disabled')
        self.personal_checkbutton.grid(row=2, column=6, sticky='w', padx=200, pady=0)
        self.to_sell_after_var = tk.StringVar()
        self.to_sell_after_label = tk.Label(self.product_frame, text='To Sell After')
        self.to_sell_after_label.grid(row=2, column=0, sticky='w', padx=0, pady=0)

        # Row 3
        self.reviewed_var = tk.BooleanVar()
        self.reviewed_checkbutton = tk.Checkbutton(self.product_frame, text='Reviewed', variable=self.reviewed_var, state='disabled')
        self.reviewed_checkbutton.grid(row=3, column=6, sticky='w', padx=200, pady=0)
        self.to_sell_after_entry = DateEntry(self.product_frame, textvariable=self.to_sell_after_var, state='disabled')
        self.to_sell_after_entry.grid(row=3, column=0, sticky='w', padx=0, pady=0)

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
        self.product_id_label.grid(row=10, column=0, sticky='w', padx=0, pady=0)
        
        self.product_id_entry = tk.Entry(self.product_frame, textvariable=self.product_id_var, state='disabled')
        self.product_id_entry.grid(row=11, column=0, sticky='w', padx=0, pady=0)

        self.product_name_var = tk.StringVar()
        self.product_name_label = tk.Label(self.product_frame, text='Product Name')
        self.product_name_label.grid(row=12, column=0, sticky='w', padx=0, pady=0)
        
        self.product_name_entry = tk.Entry(self.product_frame, textvariable=self.product_name_var, state='disabled')
        self.product_name_entry.grid(row=13, column=0, sticky='w', padx=0, pady=0)

        # Note: Product Image requires a different approach
        # This will be covered later

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

        self.product_name = self.get_folder_path_from_db(selected_product)
        
    def update_to_sell_after(self, *args):
        order_date = self.order_date_var.get()
        if order_date:
            # Convert the string to a date
            order_date = datetime.strptime(order_date, "%m/%d/%y")

            # Add 6 months to the date
            to_sell_after = order_date + relativedelta(months=6)

            # Update the to_sell_after_var
            self.to_sell_after_var.set(to_sell_after.strftime("%m/%d/%y"))


    def toggle_edit_mode(self):
        if self.edit_mode:
            self.edit_mode = False
            self.sold_checkbutton.config(state='disabled')
            # Disable all the new fields
            self.cancelled_order_checkbutton.config(state='disabled')
            self.damaged_checkbutton.config(state='disabled')
            self.personal_checkbutton.config(state='disabled')
            self.reviewed_checkbutton.config(state='disabled')
            self.pictures_downloaded_checkbutton.config(state='disabled')
            self.uploaded_to_site_checkbutton.config(state='disabled')
            self.order_date_entry.config(state='disabled')
            self.to_sell_after_entry.config(state='disabled')
            self.payment_type_combobox.config(state='disabled')
            self.asin_entry.config(state='disabled')
            self.product_id_entry.config(state='disabled')
            self.product_name_entry.config(state='disabled')
            self.product_folder_entry.config(state='disabled')
            self.fair_market_value_entry.config(state='disabled')
            self.order_details_entry.config(state='disabled')
            self.order_link_entry.config(state='disabled')
            self.sold_price_entry.config(state='disabled')
        else:
            self.edit_mode = True
            self.sold_checkbutton.config(state='normal')
            # Enable all the new fields
            self.cancelled_order_checkbutton.config(state='normal')
            self.damaged_checkbutton.config(state='normal')
            self.personal_checkbutton.config(state='normal')
            self.reviewed_checkbutton.config(state='normal')
            self.pictures_downloaded_checkbutton.config(state='normal')
            self.uploaded_to_site_checkbutton.config(state='normal')
            self.order_date_entry.config(state='normal')
            self.to_sell_after_entry.config(state='normal')
            self.payment_type_combobox.config(state='normal')
            self.asin_entry.config(state='normal')
            self.product_id_entry.config(state='normal')
            self.product_name_entry.config(state='normal')
            self.product_folder_entry.config(state='normal')
            self.fair_market_value_entry.config(state='normal')
            self.order_details_entry.config(state='normal')
            self.order_link_entry.config(state='normal')
            self.sold_price_entry.config(state='normal')

    def save(self):
        if self.sold_var.get():
            # Check if the product is already in the sold folder
            try:
                if os.path.samefile(self.product_name, os.path.join(self.sold_folder, os.path.basename(self.product_name))):
                    messagebox.showinfo("Notice", "Product is already in the Sold Folder")
                    return
            except FileNotFoundError:
                pass  # If the file doesn't exist yet, the paths can't be the same
            try:
                shutil.move(self.product_name, os.path.join(self.sold_folder, os.path.basename(self.product_name)))
                messagebox.showinfo("Success", "Moved product to Sold Folder")
                if self.folder_list.curselection():
                    self.folder_list.delete(self.folder_list.curselection())
                self.display_folders(self.folder_to_scan)  # Refresh the folder list
            except Exception as e:
                messagebox.showerror("Error", f"Error moving file: {str(e)}")
        
        # Add these lines
        self.edit_mode = True
        self.toggle_edit_mode()  # Reset edit mode
        self.edit_button.grid(row=0, column=8, sticky='w', padx=235, pady=0)  # Make sure Edit button is visible

        self.focus_search_entry()


    def get_folder_names_from_db(self):
        self.cur.execute("SELECT Folder FROM folder_paths")
        return [row[0] for row in self.cur.fetchall()]

    def get_folder_path_from_db(self, folder_name):
        self.cur.execute("SELECT Path FROM folder_paths WHERE Folder=?", (folder_name,))
        return self.cur.fetchone()[0]

    def __del__(self):
        self.conn.close()

root = tk.Tk()
root.title("Improved Inventory Manager")

# Size the window to 50% of the screen size
window_width = int(root.winfo_screenwidth() * 0.5)
window_height = int(root.winfo_screenheight() * 0.5)
root.geometry(f"{window_width}x{window_height}")

root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)
root.rowconfigure(2, weight=1)
app = Application(master=root)
root.state('zoomed')

app.mainloop()