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


def update_path_hint(label_widget, path, kind):
    """Update compact hint label showing the selected file/folder name."""
    if not path:
        label_widget.config(text="Chưa chọn", fg="#7f8fa6")
        return
    base = os.path.basename(path.rstrip("/\\"))
    prefix = "📄" if kind == "file" else "📁"
    label_widget.config(text=f"{prefix} {base}", fg="#dcdde1")


###############################################################################
# GUI
###############################################################################
root = tk.Tk()
root.title("Excel Metadata Tool")
root.geometry("820x520")
root.configure(bg="#0b1d2c")
# Window icon (optional if app.ico exists)
icon_path = os.path.join(os.path.dirname(__file__), "app.ico")
if os.path.exists(icon_path):
    try:
        root.iconbitmap(icon_path)
    except Exception:
        pass

# Palette & fonts
BG = "#0b1d2c"
CARD_BG = "#0f2536"
PANEL_BG = "#102b42"
ACCENT = "#00c4a7"
ACCENT_DARK = "#0c9c85"
TEXT_MAIN = "#e8f1f8"
TEXT_SUB = "#9fb3c8"

TITLE_FONT = ("Bahnschrift", 20, "bold")
SUBTITLE_FONT = ("Bahnschrift", 12)
LABEL_FONT = ("Bahnschrift", 11, "bold")
ENTRY_FONT = ("Bahnschrift", 11)
BUTTON_FONT = ("Bahnschrift", 11, "bold")

# ttk styling
style = ttk.Style()
style.theme_use("clam")
style.configure("Card.TFrame", background=CARD_BG)
style.configure("Accent.TButton", font=BUTTON_FONT, foreground="#0b1d2c", background=ACCENT)
style.map("Accent.TButton", background=[("active", ACCENT_DARK)], foreground=[("disabled", "#4c5a68")])
style.configure("Ghost.TButton", font=BUTTON_FONT, foreground=TEXT_MAIN, background=PANEL_BG, borderwidth=0, focusthickness=0)
style.map("Ghost.TButton", background=[("active", CARD_BG)])
style.configure(
    "Outline.TButton",
    font=BUTTON_FONT,
    foreground=TEXT_MAIN,
    background=CARD_BG,
    bordercolor=TEXT_MAIN,
    focusthickness=2,
    focuscolor=TEXT_MAIN,
    relief="solid"
)
style.map(
    "Outline.TButton",
    background=[("active", PANEL_BG)],
    foreground=[("disabled", "#7f8fa6")],
    bordercolor=[("active", TEXT_MAIN)]
)
style.configure("Success.Horizontal.TProgressbar", troughcolor=PANEL_BG, background=ACCENT, bordercolor=PANEL_BG, lightcolor=ACCENT, darkcolor=ACCENT_DARK)

settings = load_settings()

# Hero header
header_frame = tk.Frame(root, bg=PANEL_BG, padx=20, pady=16)
header_frame.grid(row=0, column=0, columnspan=4, sticky="nsew")
header_row = tk.Frame(header_frame, bg=PANEL_BG)
header_row.pack(anchor="w")
hero_icon = tk.Canvas(header_row, width=30, height=30, bg=PANEL_BG, highlightthickness=0)
hero_icon.create_oval(4, 4, 26, 26, fill=ACCENT, outline=ACCENT)
hero_icon.create_text(15, 15, text="GS", fill=PANEL_BG, font=("Bahnschrift", 9, "bold"))
hero_icon.pack(side="left", padx=(0,10))
title_lbl = tk.Label(header_row, text="Excel Metadata Tool", font=TITLE_FONT, fg=TEXT_MAIN, bg=PANEL_BG)
title_lbl.pack(side="left", anchor="w")
subtitle_lbl = tk.Label(header_frame, text="Tự động hóa xuất barcode & metadata", font=SUBTITLE_FONT, fg=TEXT_SUB, bg=PANEL_BG)
subtitle_lbl.pack(anchor="w", pady=(6,0))

# Main card
card = ttk.Frame(root, style="Card.TFrame", padding=20)
card.grid(row=1, column=0, columnspan=4, padx=24, pady=(14, 10), sticky="nsew")

# Layout helpers
card.grid_columnconfigure(1, weight=1)
card.grid_columnconfigure(3, weight=0)

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
source_file_label = tk.Label(card, text="Source File", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)
source_file_label.grid(row=0, column=0, padx=(10,10), pady=6, sticky="e")
source_file_entry = tk.Entry(card, width=48, font=ENTRY_FONT, bg="#0d2d44", fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
source_file_entry.grid(row=0, column=1, padx=6, pady=6, sticky="ew")
source_file_entry.insert(0, settings.get("source_file", ""))
ttk.Button(card, text="Browse", width=10, style="Outline.TButton", command=lambda: [browse_source_file(source_file_entry, "source_file"), update_path_hint(source_file_hint, source_file_entry.get(), "file")]).grid(row=0, column=2, padx=(8,10), pady=6, sticky="w")
source_file_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
source_file_hint.grid(row=0, column=3, padx=(6,8), sticky="w")
update_path_hint(source_file_hint, source_file_entry.get(), "file")

# Source folder row
source_folder_label = tk.Label(card, text="Source Folder", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)
source_folder_label.grid(row=1, column=0, padx=(10,10), pady=6, sticky="e")
source_folder_entry = tk.Entry(card, width=48, font=ENTRY_FONT, bg="#0d2d44", fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
source_folder_entry.grid(row=1, column=1, padx=6, pady=6, sticky="ew")
source_folder_entry.insert(0, settings.get("source_folder", ""))
ttk.Button(card, text="Browse", width=10, style="Outline.TButton", command=lambda: [browse_source_folder(source_folder_entry, "source_folder"), update_path_hint(source_folder_hint, source_folder_entry.get(), "folder")]).grid(row=1, column=2, padx=(8,10), pady=6, sticky="w")
source_folder_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
source_folder_hint.grid(row=1, column=3, padx=(6,8), sticky="w")
update_path_hint(source_folder_hint, source_folder_entry.get(), "folder")

# Template row
template_label = tk.Label(card, text="Template File", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)
template_label.grid(row=2, column=0, padx=(10,10), pady=6, sticky="e")
template_entry = tk.Entry(card, width=48, font=ENTRY_FONT, bg="#0d2d44", fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
template_entry.grid(row=2, column=1, padx=6, pady=6, sticky="ew")
template_entry.insert(0, settings.get("template", ""))
ttk.Button(card, text="Browse", width=10, style="Outline.TButton", command=lambda: [browse_file(template_entry, "template"), update_path_hint(template_hint, template_entry.get(), "file")]).grid(row=2, column=2, padx=(8,10), pady=6, sticky="w")
template_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
template_hint.grid(row=2, column=3, padx=(6,8), sticky="w")
update_path_hint(template_hint, template_entry.get(), "file")

# Output row
output_label = tk.Label(card, text="Output Folder", font=LABEL_FONT, fg=TEXT_MAIN, bg=CARD_BG)
output_label.grid(row=3, column=0, padx=(10,10), pady=6, sticky="e")
output_entry = tk.Entry(card, width=48, font=ENTRY_FONT, bg="#0d2d44", fg=TEXT_MAIN, relief="flat", insertbackground=TEXT_MAIN)
output_entry.grid(row=3, column=1, padx=6, pady=6, sticky="ew")
output_entry.insert(0, settings.get("output", ""))
ttk.Button(card, text="Browse", width=10, style="Outline.TButton", command=lambda: [browse_folder(output_entry, "output"), update_path_hint(output_hint, output_entry.get(), "folder")]).grid(row=3, column=2, padx=(8,10), pady=6, sticky="w")
output_hint = tk.Label(card, text="", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG, anchor="w")
output_hint.grid(row=3, column=3, padx=(6,8), sticky="w")
update_path_hint(output_hint, output_entry.get(), "folder")


# Progress bar centered
progress_var = tk.DoubleVar()
progress_start_time = 0
progress_frame = tk.Frame(card, bg=CARD_BG)
progress_frame.grid(row=4, column=0, columnspan=4, pady=(16, 4))
progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100, length=420, style="Success.Horizontal.TProgressbar")
progress_bar.pack(side="left", padx=(0,8))
progress_percent_label = tk.Label(progress_frame, text="0%", font=("Bahnschrift", 11, "bold"), fg=TEXT_MAIN, bg=CARD_BG)
progress_percent_label.pack(side="left")
elapsed_time_label = tk.Label(progress_frame, text="0.0s", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=CARD_BG)
elapsed_time_label.pack(side="left", padx=(10,0))


# Button frame for Convert buttons (centered)
button_frame = tk.Frame(card, bg=CARD_BG)
button_frame.grid(row=5, column=0, columnspan=4, pady=(10, 0))

ttk.Button(button_frame, text="Convert File", style="Accent.TButton", width=16, command=lambda: run_process_with_progress('file')).pack(side="left", padx=14, pady=4)
ttk.Button(button_frame, text="Convert Folder", style="Accent.TButton", width=16, command=lambda: run_process_with_progress('folder')).pack(side="left", padx=14, pady=4)

# Open Output Folder button (hidden by default)
def open_output_folder():
    folder = output_entry.get()
    if os.path.isdir(folder):
        os.startfile(folder)

open_folder_btn = ttk.Button(root, text="Open Output Folder", style="Accent.TButton", command=open_output_folder)
open_folder_btn.grid(row=2, column=0, columnspan=4, pady=(4, 0))
open_folder_btn.grid_remove()

# Footer
footer = tk.Label(root, text="Gene Solutions • Automation", font=("Bahnschrift", 10), fg=TEXT_SUB, bg=BG)
footer.grid(row=3, column=0, columnspan=4, pady=(8, 12))

root.grid_columnconfigure(0, weight=1)

# --- Progress logic ---
import threading
import time


def update_progress(percent):
    progress_var.set(percent)
    progress_percent_label.config(text=f"{int(percent)}%")
    if progress_start_time:
        elapsed = time.time() - progress_start_time
        elapsed_time_label.config(text=f"{elapsed:.1f}s")
    root.update_idletasks()



def run_process_with_progress(mode):
    progress_var.set(0)
    open_folder_btn.grid_remove()
    def task():
        global progress_start_time
        try:
            progress_start_time = time.time()
            # Simulate progress in 1% increments for smoother UI
            for i in range(0, 100):
                update_progress(i)
                time.sleep(0.05)
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
