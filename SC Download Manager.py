'''
SC Download Manager is a GUI for downloading tables from Statistics Canada
Author: Mahdi Ghafaian
Date: 2025-01-26

'''
import subprocess
import sys

# Function to install missing modules
def install_missing_modules(modules):
    for module in modules:
        try:
            __import__(module)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", module])

# List of required modules
required_modules = ['tkinter', 'requests', 'threading', 'queue', 'os', 'zipfile', 'time']

# Install missing modules
install_missing_modules(required_modules)

import tkinter as tk
from tkinter import ttk
import requests
import threading
import queue
import os
import zipfile
import time

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
tables_dir = os.path.join(script_dir, 'Tables')

# Ensure the Tables directory exists
os.makedirs(tables_dir, exist_ok=True)

download_queue = queue.Queue()
download_status = {}
table_ids_file = os.path.join(script_dir, 'Resources', 'history.txt')

def download_file(url, table_id):
    dest = os.path.join(tables_dir, f"{table_id}-eng.zip")
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()  # Raise an HTTPError for bad responses (4xx and 5xx)
        total_size = int(response.headers.get('content-length', 0))
        block_size = 8192  # Increase block size to 8 KB
        progress_bar['maximum'] = total_size

        with open(dest, 'wb') as file:
            downloaded_size = 0
            start_time = time.time()
            for data in response.iter_content(block_size):
                file.write(data)
                downloaded_size += len(data)
                progress_bar['value'] = downloaded_size
                percentage = (downloaded_size / total_size) * 100
                elapsed_time = time.time() - start_time
                speed = (downloaded_size / elapsed_time / 1024) if elapsed_time > 0 else 0  # Speed in KB/s
                if downloaded_size % (block_size * 10) == 0:  # Update GUI less frequently
                    status_label.config(text=f"Downloaded: {downloaded_size // 1024} KB / {total_size // 1024} KB ({percentage:.1f}%) / {speed:.0f} KB/s")
                    task_label.config(text=f"Downloading table {table_id}")
                    root.update_idletasks()

        download_status[table_id] = True
        update_listbox()

        if unzip_var.get():
            task_label.config(text=f"Unzipping table {table_id}")
            unzip_file(dest)
            os.remove(dest)  # Delete the zip file after unzipping

    except requests.exceptions.RequestException as e:
        status_label.config(text=f"Error downloading table {table_id}: {e}")
        task_label.config(text="Error")
        download_status[table_id] = False
        update_listbox()

def unzip_file(file_path):
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(os.path.dirname(file_path))
    except zipfile.BadZipFile as e:
        status_label.config(text=f"Error unzipping file: {e}")
        task_label.config(text="Error")

def process_queue():
    while not download_queue.empty():
        table_id = download_queue.get()
        url = f"https://www150.statcan.gc.ca/n1/tbl/csv/{table_id}-eng.zip"
        download_file(url, table_id)
    task_label.config(text="All tasks completed")
    os.startfile(tables_dir)  # Open the Tables folder

def start_download():
    table_ids = listbox.get(0, tk.END)
    for table_id in table_ids:
        if table_id not in download_status:
            download_queue.put(table_id)
            download_status[table_id] = False
    threading.Thread(target=process_queue).start()

def close_application():
    save_table_ids()
    root.destroy()

def add_table_id():
    table_id = table_id_entry.get()
    if table_id.isdigit():
        listbox.insert(tk.END, table_id)
        table_id_entry.delete(0, tk.END)
        save_table_ids()
        task_label.config(text="Table added to the download list.")

    else:
        task_label.config(text="Invalid table ID. Please enter a numeric value.")

def remove_selected_table_id():
    selected_indices = listbox.curselection()
    for index in selected_indices[::-1]:
        table_id = listbox.get(index)
        if table_id in download_status:
            del download_status[table_id]
        listbox.delete(index)
        task_label.config(text="Table removed from the download list.")
    save_table_ids()

def update_listbox():
    for i in range(listbox.size()):
        table_id = listbox.get(i)
        if download_status.get(table_id):
            listbox.itemconfig(i, {'bg':'#d3ffd3'})  # Change background color to indicate download complete
        else:
            listbox.itemconfig(i, {'bg':'white'})  # Default background color

def save_table_ids():
    with open(table_ids_file, 'w') as file:
        for table_id in listbox.get(0, tk.END):
            file.write(f"{table_id}\n")

def load_table_ids():
    if os.path.exists(table_ids_file):
        with open(table_ids_file, 'r') as file:
            for line in file:
                table_id = line.strip()
                if table_id:
                    listbox.insert(tk.END, table_id)

root = tk.Tk()
root.title("Statistics Canada Download Manager")
root.resizable(False, False)  # Freeze the window size

# Set the window icon
icon_path = os.path.join(script_dir, 'Resources', 'ca.ico')

try:
    root.iconbitmap(icon_path)
except tk.TclError:
    print(f"Icon file '{icon_path}' not found or not a valid .ico file. Please ensure the file is in the correct directory and is a valid .ico file.")

frame = ttk.Frame(root, padding=10)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

table_id_label = ttk.Label(frame, text="Table ID:")
table_id_label.grid(row=0, column=0, pady=5, sticky=(tk.W, tk.E))

# Function to handle focus in event
def on_entry_click(event):
    if table_id_entry.get() == 'Enter table ID':
        table_id_entry.delete(0, "end")  # delete all the text in the entry
        table_id_entry.config(foreground='black')

# Function to handle focus out event
def on_focus_out(event):
    if table_id_entry.get() == '':
        table_id_entry.insert(0, 'Enter table ID')
        table_id_entry.config(foreground='grey')

table_id_entry = ttk.Entry(frame, foreground='grey')
table_id_entry.insert(0, 'Enter table ID')
table_id_entry.grid(row=1, column=0, pady=5, sticky=(tk.W, tk.E))

table_id_entry.bind('<FocusIn>', on_entry_click)
table_id_entry.bind('<FocusOut>', on_focus_out)

add_button = ttk.Button(frame, text="Add", command=add_table_id)
add_button.grid(row=2, column=0, pady=5, sticky=(tk.W, tk.E))

remove_button = ttk.Button(frame, text="Remove", command=remove_selected_table_id)
remove_button.grid(row=3, column=0, pady=5, sticky=(tk.W, tk.E))

table_id_label = ttk.Label(frame, text="Download List:")
table_id_label.grid(row=4, column=0, pady=5, sticky=(tk.W, tk.E))

listbox = tk.Listbox(frame, width=table_id_entry.cget("width"))
listbox.grid(row=5, column=0, pady=5, sticky=(tk.W, tk.E))

unzip_var = tk.BooleanVar(value=True)
unzip_check = ttk.Checkbutton(frame, text="Unzip files", variable=unzip_var, width=30)
unzip_check.grid(row=6, column=0, pady=5, sticky=(tk.W, tk.E))

download_button = ttk.Button(frame, text="Download", command=start_download, width=30)
download_button.grid(row=7, column=0, pady=10, sticky=(tk.W, tk.E))

progress_bar = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=8, column=0, pady=5, sticky=(tk.W, tk.E))

status_label = ttk.Label(frame, text="Downloaded: 0 KB / 0 KB (0.0%)", width=30)
status_label.grid(row=9, column=0, pady=5, sticky=(tk.W, tk.E))

close_button = ttk.Button(frame, text="Close", command=close_application, width=30)
close_button.grid(row=10, column=0, pady=10, sticky=(tk.W, tk.E))

task_label = ttk.Label(root, text="Ready", relief=tk.FLAT, anchor=tk.W, width=30)
task_label.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)

# Set a fixed width for the window
root.update_idletasks()  # Update "requested size" from geometry manager
width = root.winfo_width()
height = root.winfo_height()
root.minsize(width, height)
root.maxsize(width, height)

load_table_ids()

root.mainloop()
