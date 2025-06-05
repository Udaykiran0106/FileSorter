import os
import shutil
from tkinter import Tk, Button, Label, filedialog, messagebox, StringVar
from pathlib import Path

# Supported file extensions
PDF_EXTENSIONS = ['.pdf']
IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.gif', '.bmp']
WORD_EXTENSIONS = ['.doc', '.docx']
EXCEL_EXTENSIONS = ['.xls', '.xlsx', '.xlsm', '.csv', '.xlsb']
PPT_EXTENSIONS = ['.ppt', '.pptx']
EXE_EXTENSIONS = ['.exe']

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        selected_folder.set(folder_path)

def sort_files():
    src_dir = selected_folder.get()
    if not src_dir or not os.path.isdir(src_dir):
        messagebox.showerror("Error", "Please select a valid folder.")
        return

    sorted_dir = os.path.join(src_dir, "Sorted Files")
    os.makedirs(sorted_dir, exist_ok=True)

    created_folders = {
        'pdf': None,
        'image': None,
        'word': None,
        'excel': None,
        'ppt': None,
        'exe': None,
        'misc': None,
    }

    for root, dirs, files in os.walk(src_dir):
        # Skip the sorted folder itself to avoid infinite loop
        if root.startswith(sorted_dir):
            continue

        for file_name in files:
            file_path = os.path.join(root, file_name)
            ext = Path(file_path).suffix.lower()

            dest_folder = None

            if ext in PDF_EXTENSIONS:
                if created_folders['pdf'] is None:
                    pdf_folder = os.path.join(sorted_dir, "PDF Files")
                    os.makedirs(pdf_folder, exist_ok=True)
                    created_folders['pdf'] = pdf_folder
                dest_folder = created_folders['pdf']

            elif ext in IMAGE_EXTENSIONS:
                if created_folders['image'] is None:
                    image_folder = os.path.join(sorted_dir, "Image Files")
                    os.makedirs(image_folder, exist_ok=True)
                    created_folders['image'] = image_folder
                dest_folder = created_folders['image']

            elif ext in WORD_EXTENSIONS:
                if created_folders['word'] is None:
                    word_folder = os.path.join(sorted_dir, "Word Files")
                    os.makedirs(word_folder, exist_ok=True)
                    created_folders['word'] = word_folder
                dest_folder = created_folders['word']

            elif ext in EXCEL_EXTENSIONS:
                if created_folders['excel'] is None:
                    excel_folder = os.path.join(sorted_dir, "Excel Files")
                    os.makedirs(excel_folder, exist_ok=True)
                    created_folders['excel'] = excel_folder
                dest_folder = created_folders['excel']

            elif ext in PPT_EXTENSIONS:
                if created_folders['ppt'] is None:
                    ppt_folder = os.path.join(sorted_dir, "PowerPoint Files")
                    os.makedirs(ppt_folder, exist_ok=True)
                    created_folders['ppt'] = ppt_folder
                dest_folder = created_folders['ppt']

            elif ext in EXE_EXTENSIONS:
                if created_folders['exe'] is None:
                    exe_folder = os.path.join(sorted_dir, "Executable Files")
                    os.makedirs(exe_folder, exist_ok=True)
                    created_folders['exe'] = exe_folder
                dest_folder = created_folders['exe']

            else:
                if created_folders['misc'] is None:
                    misc_folder = os.path.join(sorted_dir, "Miscellaneous Files")
                    os.makedirs(misc_folder, exist_ok=True)
                    created_folders['misc'] = misc_folder
                dest_folder = created_folders['misc']

            if dest_folder:
                try:
                    shutil.move(file_path, dest_folder)  # Move instead of copy
                except Exception as e:
                    print(f"Error moving {file_path}: {e}")

    messagebox.showinfo("Success", "Files sorted and moved successfully!")

# GUI Setup
app = Tk()
app.title("File Sorter")
app.geometry("400x200")

selected_folder = StringVar()

Label(app, text="Select a folder to sort:").pack(pady=10)
Button(app, text="Browse Folder", command=select_folder).pack()
Label(app, textvariable=selected_folder, wraplength=350).pack(pady=5)
Button(app, text="Sort Files", command=sort_files).pack(pady=20)

app.mainloop()
