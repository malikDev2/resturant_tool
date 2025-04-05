import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def process_file():
    input_path = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not input_path:
        return

    try:
        # Try multiple encodings
        encodings = ['utf-8', 'latin1', 'cp1252']
        for encoding in encodings:
            try:
                df = pd.read_csv(input_path, header=None, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("Could not decode file with standard encodings")

        # Rest of your processing code remains the same...
        data_start = None
        for idx, row in df.iterrows():
            if str(row[0]) == "Order Id":
                data_start = idx
                break
        
        if data_start is None:
            raise ValueError("Could not find data starting row ('Order Id')")
        
        main_data = pd.read_csv(input_path, skiprows=data_start, encoding=encoding)
        
        payouts_data = []
        for idx in range(len(df)-1, max(len(df)-10, -1), -1):
            row = df.iloc[idx]
            if not row.isnull().all():
                payouts_data.append(row)
                if len(payouts_data) >= 3:
                    break
        
        payouts_data.reverse()
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

# GUI code remains the same...
root = tk.Tk()
root.title("Vito's Report Processor")
root.geometry("400x200")

tk.Label(
    root, 
    text="Vito's Restaurant Report Processor",
    font=("Arial", 14)
).pack(pady=20)

tk.Button(
    root,
    text="Process CSV File",
    command=process_file,
    height=2,
    width=20,
    bg="#4CAF50",
    fg="white"
).pack()

tk.Label(
    root,
    text="Will create:\n1. Main order data\n2. Payouts section",
    font=("Arial", 10)
).pack(pady=20)

root.mainloop()