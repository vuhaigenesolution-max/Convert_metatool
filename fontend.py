# Source file
def browse_source_file(entry, key):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)
        save_setting(key, path)

# Source folder
def browse_source_folder(entry, key):
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)
        save_setting(key, path)
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import backend
import json
import os

def browse_file(entry, key):
    path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)
        save_setting(key, path)

def browse_folder(entry, key):
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)
        save_setting(key, path)


# Định nghĩa lại hàm run_process đúng chuẩn
def run_process():
    source = source_file_entry.get()
    template = template_entry.get()
    output = output_entry.get()

    if not source or not template or not output:
        messagebox.showerror("Error", "Please select all paths")
        return
    if not os.path.isfile(source):
        messagebox.showerror("Error", "Source must be a file when using Convert File.")
        return
    save_setting("source_file", source)
    save_setting("template", template)
    save_setting("output", output)
    try:
        backend.process_excel(source, template, output)
        messagebox.showinfo("Success", "CSV file generated successfully")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def run_process_folder():
    from backend_funtion_convert_folder import process_folder
    source = source_folder_entry.get()
    template = template_entry.get()
    output = output_entry.get()
    if not source or not template or not output:
        messagebox.showerror("Error", "Please select all paths")
        return
    if not os.path.isdir(source):
        messagebox.showerror("Error", "Source must be a folder when using Convert Folder.")
        return
    try:
        processed_files = process_folder(source, template, output)
        if processed_files:
            messagebox.showinfo("Success", f"Đã xử lý {len(processed_files)} file. Kết quả lưu tại: {output}")
        else:
            messagebox.showwarning("No Files Processed", "Không có file Excel nào được xử lý trong thư mục này.")
    except Exception as e:
        messagebox.showerror("Error", f"Lỗi khi xử lý thư mục: {e}")

def load_settings():
    settings_path = os.path.join(os.path.dirname(__file__), "settings.json")
    if os.path.exists(settings_path):
        with open(settings_path, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except Exception:
                return {"source": "", "template": "", "output": ""}
    return {"source": "", "template": "", "output": ""}

def save_setting(key, value):
    settings_path = os.path.join(os.path.dirname(__file__), "settings.json")
    settings = load_settings()
    settings[key] = value
    with open(settings_path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# ===== GUI =====
root = tk.Tk()
root.title("Excel Automation Tool")
root.geometry("700x400")
root.configure(bg="#f5f6fa")

LABEL_FONT = ("Segoe UI", 11, "bold")
ENTRY_FONT = ("Segoe UI", 11)
BUTTON_FONT = ("Segoe UI", 11, "bold")
BUTTON_COLOR = "#00a8ff"
BUTTON_FG = "#fff"

header = tk.Label(root, text="Excel Metadata Tool", font=("Segoe UI", 18, "bold"), fg="#273c75", bg="#f5f6fa")
header.grid(row=0, column=0, columnspan=4, pady=(18, 10), sticky="nsew")

settings = load_settings()

# Source (file or folder)
def browse_source(entry, key):
    # Cho phép chọn file hoặc folder
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not path:
        path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)
        save_setting(key, path)



# Source file row
source_file_label = tk.Label(root, text="Source File", font=LABEL_FONT, bg="#f5f6fa")
source_file_label.grid(row=1, column=0, padx=(30,10), pady=10, sticky="e")
source_file_entry = tk.Entry(root, width=48, font=ENTRY_FONT)
source_file_entry.grid(row=1, column=1, padx=5, pady=10, sticky="ew")
source_file_entry.insert(0, settings.get("source_file", ""))
tk.Button(root, text="Browse", font=BUTTON_FONT, bg=BUTTON_COLOR, fg=BUTTON_FG, command=lambda: browse_source_file(source_file_entry, "source_file")).grid(row=1, column=2, padx=(5,30), pady=10, sticky="w")

# Source folder row
source_folder_label = tk.Label(root, text="Source Folder", font=LABEL_FONT, bg="#f5f6fa")
source_folder_label.grid(row=2, column=0, padx=(30,10), pady=10, sticky="e")
source_folder_entry = tk.Entry(root, width=48, font=ENTRY_FONT)
source_folder_entry.grid(row=2, column=1, padx=5, pady=10, sticky="ew")
source_folder_entry.insert(0, settings.get("source_folder", ""))
tk.Button(root, text="Browse", font=BUTTON_FONT, bg=BUTTON_COLOR, fg=BUTTON_FG, command=lambda: browse_source_folder(source_folder_entry, "source_folder")).grid(row=2, column=2, padx=(5,30), pady=10, sticky="w")

# Template row
template_label = tk.Label(root, text="Template File", font=LABEL_FONT, bg="#f5f6fa")
template_label.grid(row=3, column=0, padx=(30,10), pady=10, sticky="e")
template_entry = tk.Entry(root, width=48, font=ENTRY_FONT)
template_entry.grid(row=3, column=1, padx=5, pady=10, sticky="ew")
template_entry.insert(0, settings.get("template", ""))
tk.Button(root, text="Browse", font=BUTTON_FONT, bg=BUTTON_COLOR, fg=BUTTON_FG, command=lambda: browse_file(template_entry, "template")).grid(row=3, column=2, padx=(5,30), pady=10, sticky="w")

# Output row
output_label = tk.Label(root, text="Output Folder", font=LABEL_FONT, bg="#f5f6fa")
output_label.grid(row=4, column=0, padx=(30,10), pady=10, sticky="e")
output_entry = tk.Entry(root, width=48, font=ENTRY_FONT)
output_entry.grid(row=4, column=1, padx=5, pady=10, sticky="ew")
output_entry.insert(0, settings.get("output", ""))
tk.Button(root, text="Browse", font=BUTTON_FONT, bg=BUTTON_COLOR, fg=BUTTON_FG, command=lambda: browse_folder(output_entry, "output")).grid(row=4, column=2, padx=(5,30), pady=10, sticky="w")



# Progress bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=500)
progress_bar.grid(row=5, column=0, columnspan=3, padx=30, pady=(10, 0), sticky="ew")

# Button frame for better alignment (dời xuống dòng 6)
button_frame = tk.Frame(root, bg="#f5f6fa")
button_frame.grid(row=6, column=0, columnspan=4, pady=(10, 0))

convert_file_btn = tk.Button(button_frame, text="Convert File", font=("Segoe UI", 13, "bold"), bg="#44bd32", fg="#fff", width=16, command=lambda: run_process_with_progress('file'))
convert_file_btn.pack(side="left", padx=20)

convert_folder_btn = tk.Button(button_frame, text="Convert Folder", font=("Segoe UI", 13, "bold"), bg="#273c75", fg="#fff", width=16, command=lambda: run_process_with_progress('folder'))
convert_folder_btn.pack(side="left", padx=20)

# Open Output Folder button (hidden by default)
def open_output_folder():
    import os
    import subprocess
    folder = output_entry.get()
    if os.path.isdir(folder):
        os.startfile(folder)

open_folder_btn = tk.Button(root, text="Open Output Folder", font=("Segoe UI", 11, "bold"), bg="#00a8ff", fg="#fff", command=open_output_folder)
open_folder_btn.grid(row=7, column=1, pady=(10, 0))
open_folder_btn.grid_remove()


# Dời footer xuống dòng 6

# Dời footer xuống dòng 8
footer = tk.Label(root, text="Gene Solutions - Automation", font=("Segoe UI", 10), fg="#718093", bg="#f5f6fa")
footer.grid(row=8, column=0, columnspan=4, pady=(10, 0), sticky="nsew")

root.grid_columnconfigure(1, weight=1)

# --- Progress logic ---
import threading
import time

def update_progress(percent):
    progress_var.set(percent)
    root.update_idletasks()


def run_process_with_progress(mode):
    progress_var.set(0)
    open_folder_btn.grid_remove()
    def task():
        try:
            # Simulate progress in 1% increments for smoother UI
            for i in range(0, 100):
                update_progress(i)
                time.sleep(0.01)
            # Only run the process when progress is at 99%
            if mode == 'file':
                run_process()
            elif mode == 'folder':
                run_process_folder()
            # Mark as finished only when progress is 100%
            update_progress(100)
            open_folder_btn.grid()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            update_progress(0)
    threading.Thread(target=task).start()

root.mainloop()
