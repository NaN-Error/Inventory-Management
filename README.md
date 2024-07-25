Inventory Management System


Overview
The Inventory Management System is a comprehensive solution for managing and tracking inventory. It features a user-friendly graphical user interface (GUI) built with tkinter, allowing for easy interaction and efficient inventory management.


Features
- Database Management: Utilizes the DatabaseManager class to handle all database operations using sqlite3, including database setup and folder path management.
- GUI Development: Employs tkinter for creating interactive windows, dialogs, and widgets.
- File Handling: Supports operations with different file formats (Excel, Word) using libraries like pandas, openpyxl, and python-docx.
- Date and Time Functions: Integrates datetime and dateutil modules for managing dates.
- Regular Expressions: Uses the re module for text processing, ensuring data validation and formatting.
- Image Processing: Implements PIL and openpyxl_image_loader for handling and displaying images.
- Logging: Includes logging functionality to track and record application activities, aiding in debugging and maintenance.


Requirements
Python 3.x
Required Python packages (listed in `requirements.txt`):
  - tkinter
  - pandas
  - sqlite3
  - openpyxl
  - python-docx
  - Pillow
  - tkcalendar
  - ttkthemes
 

Installation
Clone the repository:
```sh
git clone https://github.com/yourusername/inventory-management.git
cd inventory-management
```
Install dependencies:
```sh
pip install -r requirements.txt
```
Run the application:
```sh
python Inventory\ Management.py
```

Usage
Start the Application: Run the main script to start the application.
Manage Inventory: Use the GUI to add, update, delete, and view inventory items.
File Import/Export: Import and export inventory data using supported file formats (Excel, Word).
Track Dates: Utilize the date and time functions for inventory timelines.
View Logs: Check the logs for tracking and debugging application activities.


File Structure
```
Inventory-Management/
│
├── .gitignore                # Specifies files to ignore
├── requirements.txt          # Lists the required dependencies
├── Inventory Management.py   # Main script for the application
└── README.md                 # This README file
```


Author
[WB]

