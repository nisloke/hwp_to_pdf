import customtkinter as ctk

try:
    app = ctk.CTk()
    app.title("Test")
    app.geometry("200x100")
    label = ctk.CTkLabel(app, text="Hello, CustomTkinter!")
    label.pack(pady=20)
    app.mainloop()
except Exception as e:
    import traceback
    with open("error_simple.log", "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())
