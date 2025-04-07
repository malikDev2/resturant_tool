# Imports
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk  
import pandas as pd

# File Processing Code
def process_file():
    input_path = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not input_path:
        return

    try:
        encodings = ['utf-8', 'latin1', 'cp1252']
        for encoding in encodings:
            try:
                df = pd.read_csv(input_path, header=None, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("Could not decode file with standard encodings")

        data_start = None
        for idx, row in df.iterrows():
            if str(row[0]) == "Order Id":
                data_start = idx
                break
        
        if data_start is None:
            raise ValueError("Could not find data starting row ('Order Id')")
        
        main_data = pd.read_csv(input_path, skiprows=data_start, encoding=encoding)
        
        columns_to_remove = ['Order Creation Time', 'Preferred Time', 'County', 'Payment Type', 'Location']
        for col in columns_to_remove:
            if col in main_data.columns:
                main_data.drop(col, axis=1, inplace=True)
        
        last_rows = []
        for idx in range(len(df)-1, max(len(df)-10, -1), -1):
            row = df.iloc[idx]
            if not row.isnull().all():
                last_rows.append(idx)
                if len(last_rows) >= 3:
                    break
        
        if last_rows:
            main_data = main_data.iloc[:-(len(last_rows))]
        
        payouts_data = []
        for idx in last_rows:
            row = df.iloc[idx]
            payouts_data.append(row.dropna().tolist())
        
        payouts_df = pd.DataFrame(payouts_data)
        
        save_path = filedialog.asksaveasfilename(
            title="Save Main Output As",
            defaultextension=".xlsx",
            initialfile="main_output.xlsx"
        )
        if save_path:
            main_data.to_excel(save_path, index=False)
            payouts_save_path = save_path.replace(".xlsx", "_payouts.xlsx")
            payouts_df.to_excel(payouts_save_path, index=False, header=False)
            
            messagebox.showinfo(
                "Success", 
                f"Files created:\n{save_path}\n{payouts_save_path}"
            )

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed:\n{str(e)}")


# GUI Code
root = tk.Tk()
root.title("Vito's Report Processor")
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
    text="Vito's Restaurant Report Processor",
    font=("Arial", 14)
).pack(pady=5)

tk.Button(
    root,
    text="Process CSV File",
    command=process_file,
    height=2,
    width=20,
    bg="#07ebd0",
    fg="white"
).pack(pady=10)

tk.Label(
    root,
    text="Will generate:\n1. Main order data\n2. Payouts section",
    font=("Arial", 10)
).pack(pady=10)

root.mainloop()