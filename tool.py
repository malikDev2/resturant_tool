import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import re
import numpy as np

def clean_filename(name):
    """Clean the location name to be filesystem-safe"""
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()

def process_file():
    input_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if not input_path:
        return

    try:
        # Read the Excel file
        df = pd.read_excel(input_path, header=None)
        
        # Find the start of the data (row with "Order Id")
        data_start = None
        for idx, row in df.iterrows():
            if str(row[0]) == "Order Id":
                data_start = idx
                break
        
        if data_start is None:
            raise ValueError("Could not find data starting row ('Order Id')")
        
        # Read the data section with headers
        main_data = pd.read_excel(input_path, skiprows=data_start)
        
        # Clean up column names (remove extra spaces)
        main_data.columns = main_data.columns.str.strip()
        
        # Get location name from first data row
        location_name = main_data.iloc[0]['Location']
        safe_location = clean_filename(location_name)
        
        # Columns to remove
        columns_to_remove = ['Order Creation Time', 'Preferred Time', 'County', 'Payment Type', 'Location']
        for col in columns_to_remove:
            if col in main_data.columns:
                main_data.drop(col, axis=1, inplace=True)
        
        # Get the last row (original totals)
        totals_row = main_data.iloc[-1:]
        subtotal = totals_row['Subtotal'].values[0]
        ct_tax = totals_row['CT Restaurant Tax'].values[0]
        tip = totals_row['Tip'].values[0]
        delivery_fee = totals_row['Delivery Fee'].values[0]
        order_total = totals_row['Order Total'].values[0]
        
        # Remove the last row from main data
        main_data = main_data.iloc[:-1]
        
        # Calculate payable amounts
        payable_to_cc = subtotal * 0.05  # 5% of subtotal
        payable_to_doordash = tip + delivery_fee
        
        # Create the enhanced payouts dataframe
        payouts_data = pd.DataFrame({
            'Description': ['TOTALS', 'PAYABLE BREAKDOWN', '', ''],
            'Subtotal': [subtotal, payable_to_cc, np.nan, np.nan],
            'CT Restaurant Tax': [ct_tax, np.nan, ct_tax, np.nan],
            'Tip': [tip, np.nan, np.nan, payable_to_doordash],
            'Delivery Fee': [delivery_fee, np.nan, np.nan, np.nan],
            'Order Total': [order_total, np.nan, np.nan, np.nan],
            'Payable To': ['', 'Payable to CC', 'Payable to State', 'Payable to DoorDash']
        })
        
        # Format currency columns (we'll do this in Excel)
        currency_cols = ['Subtotal', 'CT Restaurant Tax', 'Tip', 'Delivery Fee', 'Order Total']
        
        save_path = filedialog.asksaveasfilename(
            title="Save Main Output As",
            defaultextension=".xlsx",
            initialfile=f"{safe_location}_main_output.xlsx"
        )
        
        if save_path:
            # Save main data
            main_data.to_excel(save_path, index=False)
            
            # Save payouts data - we'll use xlsxwriter for formatting
            payouts_save_path = save_path.replace("_main_output.xlsx", "_payouts.xlsx")
            
            # Create Excel writer object with xlsxwriter
            with pd.ExcelWriter(payouts_save_path, engine='xlsxwriter') as writer:
                payouts_data.to_excel(
                    writer,
                    index=False,
                    sheet_name='Payouts'
                )
                
                # Get the workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Payouts']
                
                # Add formatting
                header_format = workbook.add_format({
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1
                })
                
                money_format = workbook.add_format({
                    'num_format': '$#,##0.00',
                    'align': 'right'
                })
                
                # Apply header formatting
                for col_num, value in enumerate(payouts_data.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Format money columns
                money_cols = {
                    'B': 'Subtotal',
                    'C': 'CT Restaurant Tax',
                    'D': 'Tip',
                    'E': 'Delivery Fee',
                    'F': 'Order Total'
                }
                
                for col_letter, col_name in money_cols.items():
                    col_idx = payouts_data.columns.get_loc(col_name)
                    worksheet.set_column(
                        col_idx, col_idx, 15, money_format
                    )
                
                # Set other column widths
                worksheet.set_column(0, 0, 20)  # Description column
                worksheet.set_column(6, 6, 20)  # Payable To column
            
            messagebox.showinfo(
                "Success", 
                f"Files created:\n{save_path}\n{payouts_save_path}"
            )

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed:\n{str(e)}")

# GUI Code (unchanged)
root = tk.Tk()
root.title("Constant Cuisine Report Generator")
root.geometry("500x300") 

try:
    image = Image.open("logo.png")  
    image = image.resize((200, 100), Image.LANCZOS)  
    logo = ImageTk.PhotoImage(image)
    
    logo_label = tk.Label(root, image=logo)
    logo_label.image = logo  
    logo_label.pack(pady=10)
except Exception as e:
    print(f"Image not loaded: {e}")  

tk.Label(
    root, 
    text="Restaurant Report Processor",
    font=("Arial", 14)
).pack(pady=5)

tk.Button(
    root,
    text="Process Excel File",
    command=process_file,
    height=2,
    width=20,
    bg="#07ebd0",
    fg="white"
).pack(pady=10)

tk.Label(
    root,
    text="Will generate:\n1. Main order data\n2. Enhanced payouts with breakdown",
    font=("Arial", 10)
).pack(pady=10)

root.mainloop()