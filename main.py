import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from ttkbootstrap import Style
from barcode import Code39
from barcode.writer import ImageWriter
from PIL import Image, ImageTk
import win32print
from datetime import datetime
from ttkbootstrap.dialogs import Messagebox
import win32api

class BarcodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Barcode Generator")
        self.root.geometry("600x400")

        style = Style()
        style.theme_use('superhero')

        now = datetime.now()
        self.year = tk.StringVar(value=now.year)
        self.month = tk.StringVar(value=now.month)
        self.date = tk.StringVar(value=now.day)

        self.number = 0  # to start from "1" here I have to put "0"

        # Main Frame
        main_frame = ttk.Frame(root)
        main_frame.pack(expand=True, fill=tk.BOTH)

        # Project Title
        title_label = ttk.Label(main_frame, text="Barcode Generator", font=("Helvetica", 25, "bold"))
        title_label.pack(pady=10)

        # Input Frame
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(pady=20)

        ttk.Label(input_frame, text="Enter Year (YYYY):", font=("Helvetica", 10)).grid(row=0, column=0, padx=5)
        ttk.Entry(input_frame, textvariable=self.year, width=10).grid(row=0, column=1, padx=5)

        ttk.Label(input_frame, text="Select Date:").grid(row=0, column=2, padx=5)
        ttk.Combobox(input_frame, textvariable=self.date, values=list(range(1, 32)), width=5).grid(row=0, column=3, padx=5)
        ttk.Combobox(input_frame, textvariable=self.month, values=list(range(1, 13)), width=5).grid(row=0, column=4, padx=5)

        # Barcode Frame
        barcode_frame = ttk.Frame(main_frame)
        barcode_frame.pack(expand=True)

        self.canvas = tk.Canvas(barcode_frame, width=300, height=100)
        self.canvas.pack()

        # Create a SQLite database connection
        self.conn = sqlite3.connect('barcodes23june.db')
        self.cursor = self.conn.cursor()

        # Create a table to store the generated codes if it doesn't exist
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS barcodes
                               (id INTEGER PRIMARY KEY, code TEXT UNIQUE)''')
        self.conn.commit()

        # Button Frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)

        ttk.Button(button_frame, text="Generate Barcode", command=self.generate_barcode, bootstyle="success").pack(side=tk.LEFT, padx=10)
        ttk.Button(root, text="Adjust Entry", command=self.adjust_entry).pack(pady=10)
        ttk.Button(button_frame, text="Print Barcode", command=self.print_barcode, bootstyle="light").pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Previous Code", command=self.previous_code, bootstyle="info").pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Next Code", command=self.next_code, bootstyle="primary").pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Exit", command=root.destroy, bootstyle="danger").pack(side=tk.LEFT, padx=10)

        self.current_code_index = None
        self.codes = []

    def load_codes(self):
        self.cursor.execute("SELECT code FROM barcodes ORDER BY id")
        rows = self.cursor.fetchall()
        self.codes = [row[0] for row in rows]

    def generate_barcode(self):
        year_str = self.year.get()
        if not year_str.isdigit() or len(year_str) != 4:
            tk.messagebox.showerror("Input Error", "Please enter a valid 4-digit year.")
            return

        self.number += 1
        date_str = self.date.get()
        month_str = self.month.get()
        if not (date_str and month_str):
            tk.messagebox.showerror("Input Error", "Please select a complete date.")
            return

        date_month_year = f"{year_str[-2:]}{month_str.zfill(2)}{date_str.zfill(2)}"
        code_str = date_month_year + f"{self.number:04d}"

        if self.code_exists(code_str):
            tk.messagebox.showwarning("Duplicate Code", "This code already exists. Please press the Adjust Entry button.")
            self.number -= 1  # Decrement the number to avoid skipping numbers
        else:
            self.update_barcode(date_month_year)

    def adjust_entry(self):
        current_date = datetime.now().strftime('%y%m%d')
        date_month_year = f"{current_date}"

        # Get the last serial number from the database
        self.cursor.execute("SELECT MAX(code) FROM barcodes")
        last_serial = self.cursor.fetchone()[0]

        if last_serial is not None:
            self.number = int(last_serial[-4:]) + 1
        else:
            self.number = 1

        self.update_barcode(date_month_year)

    def update_barcode(self, date_month_year):
        number_str = f"{self.number:04d}"  # Create a 4-digit number with leading zeros
        code_str = date_month_year + number_str  # Combine date, month, year, and number

        my_code = Code39(code_str, writer=ImageWriter(), add_checksum=False)
        my_code.save("barcode")

        image = Image.open("barcode.png")
        image = image.resize((300, 100), Image.LANCZOS)
        self.barcode_image = ImageTk.PhotoImage(image)

        # Insert the generated code into the database
        self.cursor.execute("INSERT INTO barcodes (code) VALUES (?)", (code_str,))
        self.conn.commit()

        self.canvas.create_image(150, 50, image=self.barcode_image)

        self.load_codes()
        self.current_code_index = len(self.codes) - 1

    def code_exists(self, code_str):
        self.cursor.execute("SELECT COUNT(*) FROM barcodes WHERE code = ?", (code_str,))
        count = self.cursor.fetchone()[0]
        return count > 0

    def display_code(self, code_str):
        my_code = Code39(code_str, writer=ImageWriter(), add_checksum=False)
        my_code.save("barcode")

        image = Image.open("barcode.png")
        image = image.resize((300, 100), Image.LANCZOS)
        self.barcode_image = ImageTk.PhotoImage(image)

        self.canvas.create_image(150, 50, image=self.barcode_image)

    def previous_code(self):
        if self.current_code_index is None:
            self.load_codes()
            self.current_code_index = len(self.codes) - 1

        if self.current_code_index > 0:
            self.current_code_index -= 1
            self.display_code(self.codes[self.current_code_index])

    def next_code(self):
        if self.current_code_index is None:
            self.load_codes()
            self.current_code_index = len(self.codes) - 1

        if self.current_code_index < len(self.codes) - 1:
            self.current_code_index += 1
            self.display_code(self.codes[self.current_code_index])

    def get_all_barcodes(self):
        self.cursor.execute("SELECT code FROM barcodes")
        rows = self.cursor.fetchall()
        return [row[0] for row in rows]

    def print_barcode(self):
        printer_name = win32print.GetDefaultPrinter()
        hdc = win32print.CreateDC("WINSPOOL", printer_name, None)
        hdc.StartDoc("Barcode Print")
        hdc.StartPage()

        # Scale to fit 3-inch width paper (300 DPI)
        image = Image.open("barcode.png")
        width, height = image.size
        aspect_ratio = height / width
        width_inch = 3
        height_inch = width_inch * aspect_ratio

        # Convert inches to pixels (300 DPI)
        width_px = int(width_inch * 300)
        height_px = int(height_inch * 300)

        # Resize image for printing
        image = image.resize((width_px, height_px), Image.LANCZOS)
        image.save("barcode_print.bmp")

        # Print the image
        win32print.SetGraphicsMode(hdc, win32print.GM_ADVANCED)
        bmp = Image.open("barcode_print.bmp")
        bmp_bits = bmp.tobytes()

        dib = win32print.CreateDIBitmap(hdc, bmp_bits)
        win32print.SetDIBitsToDevice(
            hdc, 0, 0, width_px, height_px, 0, 0, 0, height_px, bmp_bits
        )

        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()

if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeApp(root)
    root.eval('tk::PlaceWindow %s center' % root.winfo_pathname(root.winfo_id()))
    root.mainloop()
