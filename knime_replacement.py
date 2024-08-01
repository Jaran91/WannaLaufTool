import os
import email
from email.parser import Parser
import pandas as pd
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog, messagebox

def read_eml_files(directory):
    data_list = []
    for filename in os.listdir(directory):
        if filename.endswith(".eml"):
            file_path = os.path.join(directory, filename)
            with open(file_path, "rb") as file:
                msg = email.message_from_binary_file(file)
                html_content = msg.get_payload(decode=True)
                soup = BeautifulSoup(html_content, 'html.parser')
                        
                        # Extract data from HTML
                data = {}
                rows = soup.find_all('tr')
                for row in rows:
                    cells = row.find_all('td')
                    if len(cells) == 2:                            
                        key = cells[0].get_text(strip=True)
                        value = cells[1].get_text(strip=True)
                        data[key] = value
                        
                data_list.append(data)
    data_list = [dict(t) for t in {tuple(d.items()) for d in data_list}]
    data_list = sorted(data_list, key=lambda x: x.get('Nachname', ''))

                       
    return data_list
                       
def export_to_excel(data_list, output_path):
    df = pd.DataFrame(data_list)
    df.to_excel(output_path, index=False)

def select_directory():
    directory = filedialog.askdirectory()
    if directory:
        directory_path.set(directory)
        print(f"Selected directory: {directory}")  # Debug print

def open_excel_file():
    output_path = excel_path.get()
    if output_path:
        os.startfile(output_path)

def start_process():
    directory = directory_path.get()
    if not directory:
        messagebox.showerror("Error", "Please select a directory first")
        return
    data_list = read_eml_files(directory)
    if data_list:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            export_to_excel(data_list, output_path)
            excel_path.set(output_path)
            open_button.config(state=tk.NORMAL)
            messagebox.showinfo("Success", f"Processed {len(data_list)} files and exported to {output_path}")
    else:
        messagebox.showinfo("No Data", "No .eml files found in the selected directory")

if __name__ == '__main__':
    root = tk.Tk()
    root.title("EML File Processor")

    directory_path = tk.StringVar()
    excel_path = tk.StringVar()

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    select_button = tk.Button(frame, text="Select Directory", command=select_directory)
    select_button.pack(side=tk.LEFT, padx=5)

    start_button = tk.Button(frame, text="Start", command=start_process)
    start_button.pack(side=tk.LEFT, padx=5)

    directory_label = tk.Label(frame, textvariable=directory_path)
    directory_label.pack(side=tk.LEFT, padx=5)

    excel_label = tk.Label(root, textvariable=excel_path)
    excel_label.pack(pady=5)

    open_button = tk.Button(root, text="Open Excel File", command=open_excel_file, state=tk.DISABLED)
    open_button.pack(pady=5)

    root.mainloop()