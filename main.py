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

        # 상단 안내 레이블
        self.info_label = ctk.CTkLabel(
            self.main_frame, 
            text="💡 팁: HWP 파일을 'HwpToPdfConverter.exe'에 직접 드래그하면 자동 변환됩니다!",
            font=("", 11),
            text_color="gray"
        )
        self.info_label.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")

        self.top_frame = ctk.CTkFrame(self.main_frame)
        self.top_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        
        # 버튼을 더 크고 명확하게
        self.btn_select_files = ctk.CTkButton(
            self.top_frame, 
            text="📄 파일 선택", 
            command=self.select_files,
            height=40,
            font=("", 13, "bold")
        )
        self.btn_select_files.pack(side="left", padx=5, pady=5, fill="x", expand=True)
        
        self.btn_select_folder = ctk.CTkButton(
            self.top_frame, 
            text="📁 폴더 선택", 
            command=self.select_folder,
            height=40,
            font=("", 13, "bold")
        )
        self.btn_select_folder.pack(side="left", padx=5, pady=5, fill="x", expand=True)

        self.btn_clear_list = ctk.CTkButton(
            self.top_frame, 
            text="🗑️ 목록 지우기", 
            command=self.clear_file_list,
            height=40,
            fg_color="gray40",
            hover_color="gray50"
        )
        self.btn_clear_list.pack(side="left", padx=5, pady=5)

        self.btn_convert = ctk.CTkButton(
            self.top_frame, 
            text="⚡ 변환 시작", 
            state="disabled", 
            command=self.start_conversion,
            height=40,
            font=("", 14, "bold"),
            fg_color="green",
            hover_color="darkgreen"
        )
        self.btn_convert.pack(side="right", padx=5, pady=5)

        self.scrollable_frame = ctk.CTkScrollableFrame(
            self.main_frame, 
            label_text="📋 변환할 파일 목록",
            label_font=("", 13, "bold")
        )
        self.scrollable_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            self.main_frame, 
            text="버튼을 클릭하여 파일을 선택하세요.", 
            anchor="w",
            font=("", 12)
        )
        self.status_label.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

        # 시작 시 커맨드라인 인자 처리
        self.after(100, self.process_command_line_args)

    def process_command_line_args(self):
        """커맨드라인 인자로 전달된 파일/폴더 처리"""
        if len(sys.argv) > 1:
            for arg in sys.argv[1:]:
                if os.path.exists(arg):
                    if arg.lower().endswith(('.hwp', '.hwpx')):
                        if arg not in self.file_list:
                            self.add_file_to_list(arg)
                    elif os.path.isdir(arg):
                        self.add_files_from_folder(arg)
            
            if self.file_list:
                self.update_ui_states()
                # 자동으로 변환 시작
                self.after(500, self.start_conversion)

    def select_files(self):
        """파일 선택 대화상자"""
        file_paths = filedialog.askopenfilenames(
            title="HWP 파일 선택",
            filetypes=[
                ("HWP 파일", "*.hwp *.hwpx"),
                ("모든 파일", "*.*")
            ]
        )
        
        if not file_paths:
            return
        
        added_count = 0
        for file_path in file_paths:
            if file_path not in self.file_list:
                self.add_file_to_list(file_path)
                added_count += 1
        
        if added_count > 0:
            self.status_label.configure(text=f"✅ {added_count}개의 파일이 추가되었습니다.")
        
        self.update_ui_states()

    def select_folder(self):
        """폴더 선택 대화상자"""
        folder_path = filedialog.askdirectory(title="폴더 선택 - HWP 파일 검색")
        if not folder_path:
            return

        added_count = self.add_files_from_folder(folder_path)
        
        if added_count > 0:
            self.status_label.configure(text=f"✅ {added_count}개의 파일이 추가되었습니다.")
        else:
            self.status_label.configure(text="❌ HWP 파일을 찾을 수 없습니다.")
        
        self.update_ui_states()

    def add_files_from_folder(self, folder_path):
        """폴더에서 HWP 파일 추가"""
        added_count = 0
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(('.hwp', '.hwpx')):
                    full_path = os.path.join(root, file)
                    if full_path not in self.file_list:
                        self.add_file_to_list(full_path)
                        added_count += 1
        return added_count

    def add_file_to_list(self, file_path):
        """파일 목록에 추가"""
        self.file_list.append(file_path)
        
        # 파일명만 표시하되, 전체 경로는 툴팁으로
        file_name = os.path.basename(file_path)
        folder_name = os.path.basename(os.path.dirname(file_path))
        display_text = f"{file_name}  📂 ({folder_name})"
        
        checkbox = ctk.CTkCheckBox(
            self.scrollable_frame, 
            text=display_text,
            font=("", 11)
        )
        checkbox.full_path = file_path
        checkbox.grid(sticky="w", padx=10, pady=2)
        checkbox.select()

    def clear_file_list(self):
        """파일 목록 초기화"""
        self.file_list.clear()
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.update_ui_states()

    def update_ui_states(self, is_converting=False):
        """UI 상태 업데이트"""
        if is_converting:
            self.btn_select_folder.configure(state="disabled")
            self.btn_select_files.configure(state="disabled")
            self.btn_convert.configure(state="disabled")
            self.btn_clear_list.configure(state="disabled")
        else:
            self.btn_select_folder.configure(state="normal")
            self.btn_select_files.configure(state="normal")
            self.btn_clear_list.configure(state="normal")
            if self.file_list:
                self.btn_convert.configure(state="normal")
                self.status_label.configure(text=f"📊 총 {len(self.file_list)}개의 파일이 준비되었습니다.")
            else:
                self.btn_convert.configure(state="disabled")
                self.status_label.configure(text="버튼을 클릭하여 파일을 선택하세요.")

    def start_conversion(self):
        """변환 시작"""
        selected_files = []
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, ctk.CTkCheckBox) and widget.get() == 1:
                selected_files.append(widget.full_path)
        
        if not selected_files:
            self.status_label.configure(text="⚠️ 변환할 파일을 선택하세요.")
            return
        
        self.update_ui_states(is_converting=True)
        
        conversion_thread = threading.Thread(
            target=self.run_conversion, 
            args=(selected_files,),
            daemon=True
        )
        conversion_thread.start()

    def run_conversion(self, files_to_convert):
        """변환 실행 (별도 스레드)"""
        from converter import convert_to_pdf
        total_files = len(files_to_convert)
        success_count = 0
        
        for i, file_path in enumerate(files_to_convert):
            base_name = os.path.basename(file_path)
            
            self.after(0, self.update_status_safe, 
                      f"⏳ ({i+1}/{total_files}) 변환 중: {base_name}")
            self.after(0, self.update_progress_safe, (i + 1) / total_files * 0.9)

            if convert_to_pdf(file_path):
                success_count += 1
        
        self.after(0, self.update_progress_safe, 1.0)
        
        if success_count == total_files:
            status_msg = f"✅ 변환 완료! 총 {total_files}개 파일 모두 성공"
        else:
            status_msg = f"⚠️ 변환 완료: 총 {total_files}개 중 {success_count}개 성공"
        
        self.after(0, self.update_status_safe, status_msg)
        self.after(0, self.update_ui_states, False)

    def update_status_safe(self, message):
        """스레드 안전 상태 업데이트"""
        self.status_label.configure(text=message)

    def update_progress_safe(self, value):
        """스레드 안전 진행률 업데이트"""
        self.progress_bar.set(value)

if __name__ == "__main__":
    try:
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
            
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        app = App()
        app.mainloop()
    except Exception as e:
        import traceback
        with open("error.log", "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        raise
    finally:
        pythoncom.CoUninitialize()
