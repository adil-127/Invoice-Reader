import os
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_values_from_invoice(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[1]
        text = page.extract_text()

        amount_collected_pattern = "Amount collected by Foodpanda"
        total_commission_pattern = "Total Commission [2]"
        total_amount_pattern = "Total Amount to be paid by Foodpanda [1+4]"
        date_pattern = 'Invoice Date'
        branch_pattern = 'Partner Code'

        amount_collected_index = text.find(amount_collected_pattern)
        total_commission_index = text.find(total_commission_pattern)
        total_amount_index = text.find(total_amount_pattern)
        date_index = text.find(date_pattern)
        branch_index = text.find(branch_pattern)

        amount_collected = text[amount_collected_index:].split("\n")[0].split('-')[1].strip()
        total_commission = text[total_commission_index:].split("\n")[0].strip().split()[-1]
        total_amount = text[total_amount_index:].split("\n")[0].split("-")[1].strip()
        date = text[date_index:].split("\n")[0].split(':')[1].strip().split()[0]
        branch = text[branch_index:].split('\n')[0].split(':')[1].strip().split()[0]

        return amount_collected, total_commission, total_amount, date, branch

def process_invoices():
    folder_path = folder_path_entry.get()
    if not os.path.isdir(folder_path):
        messagebox.showerror("Error", "Invalid folder path.")
        return

    data_list = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.pdf'):
            invoice_path = os.path.join(folder_path, file_name)
            amount_collected, total_commission, total_amount, date, branch = extract_values_from_invoice(invoice_path)
            data_list.append([branch, amount_collected, total_commission, total_amount, date])

    data = pd.DataFrame(data_list, columns=['Branch', 'Amount Collected by foodpanda', 'Total Commission', 'Total Amount paid by foodpanda', 'Date'])
    excel_file_path = 'invoice_data.xlsx'
    data.to_excel(excel_file_path, index=False)
    messagebox.showinfo("Success", f"Data saved to: {excel_file_path}")

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(tk.END, folder_path)

root = tk.Tk()
root.title("Invoice Processing")

root.config(bg="#2C2F33")
root.option_add("*TEntry*background", "#40444B")
root.option_add("*TEntry*foreground", "#FFFFFF")
root.option_add("*TEntry*insertbackground", "#FFFFFF")
root.option_add("*TButton*background", "#7289DA")
root.option_add("*TButton*foreground", "#FFFFFF")
root.option_add("*TLabel*background", "#2C2F33")
root.option_add("*TLabel*foreground", "#FFFFFF")

folder_path_label = tk.Label(root, text="Folder Path:")
folder_path_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
folder_path_entry = tk.Entry(root, width=50)
folder_path_entry.grid(row=0, column=1, padx=5, pady=5)
browse_button = tk.Button(root, text="Browse", command=browse_folder)
browse_button.grid(row=0, column=2, padx=5, pady=5)

process_button = tk.Button(root, text="Process Invoices", command=process_invoices)
process_button.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
