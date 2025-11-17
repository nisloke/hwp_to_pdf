import pythoncom
# GUI의 파일 대화상자와 같은 기능이 OLE/COM을 사용하므로,
# 다른 라이브러리보다 먼저 올바른 스레드 모델로 COM을 초기화합니다.
pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

import customtkinter as ctk
import os
from customtkinter import filedialog
import threading
import sys

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("HWP to PDF 변환기")
        self.geometry("800x600")

        self.file_list = []

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        self.top_frame = ctk.CTkFrame(self.main_frame)
        self.top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        self.btn_select_folder = ctk.CTkButton(self.top_frame, text="폴더 선택", command=self.select_folder)
        self.btn_select_folder.pack(side="left", padx=5, pady=5)

        self.btn_convert = ctk.CTkButton(self.top_frame, text="선택 파일 변환", state="disabled", command=self.start_conversion)
        self.btn_convert.pack(side="right", padx=5, pady=5)
        
        self.btn_clear_list = ctk.CTkButton(self.top_frame, text="목록 지우기", command=self.clear_file_list)
        self.btn_clear_list.pack(side="right", padx=5, pady=5)

        self.scrollable_frame = ctk.CTkScrollableFrame(self.main_frame, label_text="변환할 파일 목록")
        self.scrollable_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(self.main_frame, text="폴더를 선택하여 변환할 파일을 추가하세요.", anchor="w")
        self.status_label.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        self.after(100, self.process_command_line_args)

    def process_command_line_args(self):
        if len(sys.argv) > 1:
            file_path = sys.argv[1]
            if os.path.exists(file_path) and file_path.lower().endswith(('.hwp', '.hwpx')):
                self.add_file_to_list(file_path)
                self.update_ui_states()

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if not folder_path:
            return

        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(('.hwp', '.hwpx')):
                    full_path = os.path.join(root, file)
                    if full_path not in self.file_list:
                        self.add_file_to_list(full_path)
        
        self.update_ui_states()

    def add_file_to_list(self, file_path):
        self.file_list.append(file_path)
        
        checkbox = ctk.CTkCheckBox(self.scrollable_frame, text=os.path.basename(file_path))
        checkbox.full_path = file_path
        checkbox.grid(sticky="w", padx=10, pady=2)
        checkbox.select()

    def clear_file_list(self):
        self.file_list.clear()
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.update_ui_states()

    def update_ui_states(self, is_converting=False):
        if is_converting:
            self.btn_select_folder.configure(state="disabled")
            self.btn_convert.configure(state="disabled")
            self.btn_clear_list.configure(state="disabled")
        else:
            self.btn_select_folder.configure(state="normal")
            self.btn_clear_list.configure(state="normal")
            if self.file_list:
                self.btn_convert.configure(state="normal")
                self.status_label.configure(text=f"{len(self.file_list)}개의 파일이 추가되었습니다.")
            else:
                self.btn_convert.configure(state="disabled")
                self.status_label.configure(text="폴더를 선택하여 변환할 파일을 추가하세요.")

    def start_conversion(self):
        selected_files = []
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, ctk.CTkCheckBox) and widget.get() == 1:
                selected_files.append(widget.full_path)
        
        if not selected_files:
            self.status_label.configure(text="변환할 파일을 선택하세요.")
            return
        
        self.update_ui_states(is_converting=True)
        
        conversion_thread = threading.Thread(target=self.run_conversion, args=(selected_files,))
        conversion_thread.start()

    def run_conversion(self, files_to_convert):
        from converter import convert_to_pdf
        total_files = len(files_to_convert)
        success_count = 0
        
        for i, file_path in enumerate(files_to_convert):
            base_name = os.path.basename(file_path)
            
            self.after(0, self.update_status_safe, f"({i+1}/{total_files}) 변환 중: {base_name}")
            self.after(0, self.update_progress_safe, (i + 1) / total_files * 0.9)

            if convert_to_pdf(file_path):
                success_count += 1
        
        self.after(0, self.update_progress_safe, 1.0)
        self.after(0, self.update_status_safe, f"변환 완료: 총 {total_files}개 중 {success_count}개 성공")
        self.after(0, self.update_ui_states, False)

    def update_status_safe(self, message):
        self.status_label.configure(text=message)

    def update_progress_safe(self, value):
        self.progress_bar.set(value)

if __name__ == "__main__":
    try:
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except ImportError:
            pass
            
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        app = App()
        app.mainloop()
    except Exception as e:
        import traceback
        with open("error.log", "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
    finally:
        pythoncom.CoUninitialize()