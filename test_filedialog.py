import pythoncom
pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

import customtkinter as ctk
from customtkinter import filedialog
import sys

def open_dialog():
    print("Opening file dialog...")
    path = filedialog.askdirectory()
    print(f"Dialog closed. Selected path: {path}")

try:
    app = ctk.CTk()
    app.title("FileDialog Test")
    app.geometry("300x150")
    
    button = ctk.CTkButton(app, text="Select Folder", command=open_dialog)
    button.pack(pady=40)
    
    app.mainloop()

except Exception as e:
    import traceback
    with open("error_dialog.log", "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())
finally:
    pythoncom.CoUninitialize()
