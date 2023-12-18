import tkinter as tk
from ctypes import windll
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess
import os
import math
import threading
import re
import time

# Global process to hold the ongoing compression process
process = None

def get_folder_size(folder_path):
    """Returns the total size of all the files in the folder."""
    total_size = 0
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            total_size += os.path.getsize(file_path)
    return total_size

def set_taskbar_progress(progress):
    """Set progress on taskbar."""
    try:
        windll.shcore.SetTaskbarProgressValue(root_frame.winfo_id(), progress, 100)  
        if progress == 100:
            windll.shcore.SetTaskbarProgressState(root_frame.winfo_id(), 3)  # Set state to "Done"
        else:
            windll.shcore.SetTaskbarProgressState(root_frame.winfo_id(), 2)  # Set state to "Running"
    except Exception as e:
        print(f"Unable to set taskbar progress. Exception: {e}")

def convert_size(size_bytes):
    """Converts size in bytes to a more readable format."""
    if size_bytes == 0:
        return "0 Bytes"
    size_name = ("Bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_name[i]}"

def calculate_size_reduction(original_size, compressed_size):
    """Calculates the size reduction percentage and the amount saved."""
    if original_size == 0:
        return "Error: Original size is 0", "0 Bytes"
    reduction = ((original_size - compressed_size) / original_size) * 100
    saved = original_size - compressed_size
    return f"{reduction:.2f}%", convert_size(saved)

# Replace the update_progress function with the following:
def update_progress(progress_label, console, process, input_folder, output_file, total_size_kb, progress_bar):
    """Update the progress of the compression in the GUI."""
    global start_time
    start_time = time.time()  # Start time for calculating elapsed time
    while True:
        output = process.stdout.readline()
        if output == '':
            if process.poll() is not None:
                # Process completed
                if os.path.exists(output_file):
                    elapsed_time = time.time() - start_time
                    progress_label.config(text=f"Done in {elapsed_time:.2f} seconds")
                    progress_bar['value'] = 100  # Update progress bar to 100%
                    set_taskbar_progress(100)  # Update taskbar progress to 100%
                    original_size = get_folder_size(input_folder)
                    compressed_size = os.path.getsize(output_file)
                    reduction, saved = calculate_size_reduction(original_size, compressed_size)
                    results_var.set(f"Reduction: {reduction} | Saved: {saved}")
                else:
                    results_var.set("Error: Compression failed or output file not found.")
                break
        else:
            console.config(state='normal')  # Enable the console to insert text
            console.insert(tk.END, output)
            console.config(state='disabled')  # Disable the console to prevent editing
            console.yview(tk.END)  # Auto-scroll to the end
            match = re.search(r'(\d+)KB -> (\d+)KB (\d+)KB/s', output.strip())
            if match:
                original_kb, current_kb, speed_kb = map(int, match.groups())
                remaining_kb = total_size_kb - current_kb
                if speed_kb > 0:
                    time_left = remaining_kb / speed_kb
                    time_left_str = f"{time_left:.2f} seconds" if time_left < 60 else f"{time_left/60:.2f} minutes"
                    progress_label.config(text=f"Time remaining: {time_left_str}")
                    progress_percentage = round((current_kb / total_size_kb) * 100)
                    progress_bar['value'] = progress_percentage  # Update progress bar
                    set_taskbar_progress(progress_percentage)  # Update taskbar progress
            root.update_idletasks()

def run_compression(input_folder, output_file, progress_label, console, total_size_kb, progress_bar):
    """Run the MCM compression process."""
    global process

    command = f"mcm.exe -x11 \"{input_folder}\" \"{output_file}\""
    process = subprocess.Popen(
        command, 
        stdout=subprocess.PIPE, 
        stderr=subprocess.STDOUT, 
        universal_newlines=True, 
        shell=True,
    )
    threading.Thread(target=update_progress, args=(progress_label, console, process, input_folder, output_file, total_size_kb, progress_bar)).start()

def stop_compression():
    """Stop the MCM compression process."""
    global process
    if process:
        process.terminate()

def compress_folder(input_folder, progress_label, console):
    """Start the folder compression process."""
    global total_size_kb
    if not input_folder:
        input_folder = filedialog.askdirectory()

    if input_folder:
        total_size_kb = get_folder_size(input_folder) // 1024
        input_entry.delete(0, tk.END)
        input_entry.insert(0, input_folder)
        output_file = os.path.join(os.path.dirname(input_folder), os.path.basename(input_folder) + ".mcm")
        output_entry.delete(0, tk.END)
        output_entry.insert(0, output_file)
        progress_bar = ttk.Progressbar(root, length=500)  # Change length as needed
        progress_bar.pack(padx=5, pady=5)
        run_compression(input_folder, output_file, progress_label, console, total_size_kb, progress_bar)

def on_drop(event):
    """Called when a folder is dragged and dropped onto the window."""
    files = root.tk.splitlist(event.data)
    if files:
        if os.path.isdir(files[0]):
            input_entry.delete(0, tk.END)
            input_entry.insert(0, files[0])
            output_file = os.path.join(os.path.dirname(files[0]), os.path.basename(files[0]) + ".mcm")
            output_entry.delete(0, tk.END)
            output_entry.insert(0, output_file)
        else:
            messagebox.showerror("Error", "Please drop a folder, not a file.")

root = TkinterDnD.Tk()
root.title("MCM Compressor")

input_entry = tk.Entry(root, width=50)
input_entry.pack(padx=5, pady=5)

output_entry = tk.Entry(root, width=50)
output_entry.pack(padx=5, pady=5)

results_var = tk.StringVar()
results_label = tk.Label(root, textvariable=results_var)
results_label.pack(padx=5, pady=5)

progress_label = tk.Label(root, text="Time remaining: calculating...")
progress_label.pack(padx=5, pady=5)

console = scrolledtext.ScrolledText(root, state='disabled', height=8)
console.pack(padx=5, pady=5)

button_frame = tk.Frame(root)
button_frame.pack(padx=5, pady=20)
compress_button = tk.Button(
    button_frame,
    text="Compress Folder",
    command=lambda: compress_folder(input_entry.get(), progress_label, console)
)
compress_button.pack(side=tk.LEFT, padx=5, ipadx=10)
stop_button = tk.Button(button_frame, text="Stop Compression", command=stop_compression)
stop_button.pack(side=tk.LEFT, padx=5, ipadx=10)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)

root.mainloop()