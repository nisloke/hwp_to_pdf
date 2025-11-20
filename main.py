import pythoncom
pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

import customtkinter as ctk
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, simpledialog
import os
import time
import threading
import sys
import subprocess
from PIL import Image, ImageTk, ImageDraw
from customtkinter import filedialog

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("HWP to PDF 변환기 Pro")
        self.geometry("1000x750")
        self.is_running = True

        # 이미지 리소스 초기화 (체크박스용)
        self.init_images()

        # 스타일 설정
        self.setup_styles()

        # 레이아웃 구성
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1) # PanedWindow 영역

        # 1. 상단 팁 레이블
        self.info_label = ctk.CTkLabel(
            self.main_frame, 
            text="💡 팁: 헤더를 우클릭하여 너비를 자동 조절할 수 있습니다. 결과창에서 우클릭하여 파일 관리가 가능합니다.",
            font=("", 12),
            text_color="gray"
        )
        self.info_label.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")

        # 2. 상단 버튼 영역
        self.top_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.top_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.create_buttons()

        # 3. 메인 리스트 영역 (PanedWindow로 상하 분리)
        # ttk.PanedWindow 스타일링이 제한적이므로 배경색을 다크 테마에 맞춤
        self.paned_window = ttk.PanedWindow(self.main_frame, orient="vertical")
        self.paned_window.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

        # [상단] 입력 파일 목록
        self.input_frame = ctk.CTkFrame(self.paned_window)
        self.create_input_tree(self.input_frame)
        self.paned_window.add(self.input_frame, weight=3) # 상단이 기본적으로 더 크게

        # [하단] 결과 파일 목록
        self.result_frame = ctk.CTkFrame(self.paned_window)
        self.create_result_tree(self.result_frame)
        self.paned_window.add(self.result_frame, weight=1)

        # 4. 하단 진행률 및 상태
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            self.main_frame, 
            text="파일을 추가하여 목록을 구성하세요.", 
            anchor="w",
            font=("", 12)
        )
        self.status_label.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

        # 공통 헤더 메뉴
        self.header_menu = tk.Menu(self, tearoff=0)
        self.header_menu.add_command(label="↔ 이 열 너비 맞춤", command=self.autosize_current_col)
        self.header_menu.add_command(label="↔ 모든 열 너비 맞춤", command=self.autosize_all_cols)
        self.current_header_col = None
        self.current_header_tree = None

        # 초기화
        self.after(100, self.process_command_line_args)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("default")
        
        # Treeview 공통 스타일
        style.configure("Treeview", 
                        rowheight=26, 
                        font=("", 10),
                        background="#2b2b2b", 
                        foreground="white",
                        fieldbackground="#2b2b2b",
                        borderwidth=0)
        style.configure("Treeview.Heading", 
                        font=("", 10, "bold"),
                        background="#3a3a3a", 
                        foreground="white",
                        relief="flat")
        style.map("Treeview", background=[("selected", "#1f538d")])
        style.map("Treeview.Heading", background=[("active", "#4a4a4a")])
        
        # PanedWindow 새시(Sash) 스타일
        style.configure("Sash", background="#1f538d", sashthickness=5)

    def init_images(self):
        """체크박스 이미지 생성 (Pillow 사용)"""
        size = (16, 16)
        
        # 1. 체크 안 된 상태 (빈 박스)
        self.img_unchecked_pil = Image.new("RGBA", size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(self.img_unchecked_pil)
        draw.rectangle([1, 1, 14, 14], outline="#aaaaaa", width=1)
        self.img_unchecked = ImageTk.PhotoImage(self.img_unchecked_pil)

        # 2. 체크 된 상태 (파란 박스 + 체크)
        self.img_checked_pil = Image.new("RGBA", size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(self.img_checked_pil)
        draw.rectangle([1, 1, 14, 14], fill="#1f538d", outline="#1f538d")
        # 체크 표시 (하얀색)
        draw.line([3, 8, 6, 11], fill="white", width=2)
        draw.line([6, 11, 12, 4], fill="white", width=2)
        self.img_checked = ImageTk.PhotoImage(self.img_checked_pil)

    def create_buttons(self):
        btn_params = {"height": 35, "width": 100, "font": ("", 13, "bold")}
        
        self.btn_files = ctk.CTkButton(self.top_frame, text="📄 파일 추가", command=self.select_files, **btn_params)
        self.btn_files.pack(side="left", padx=5)
        
        self.btn_folder = ctk.CTkButton(self.top_frame, text="📁 폴더 추가", command=self.select_folder, **btn_params)
        self.btn_folder.pack(side="left", padx=5)

        self.btn_remove = ctk.CTkButton(self.top_frame, text="❌ 체크 제거", command=self.remove_checked_files, 
                                        fg_color="gray40", hover_color="gray50", **btn_params)
        self.btn_remove.pack(side="left", padx=5)

        self.btn_clear = ctk.CTkButton(self.top_frame, text="🗑️ 전체 삭제", command=self.clear_file_list, 
                                       fg_color="gray40", hover_color="gray50", **btn_params)
        self.btn_clear.pack(side="left", padx=5)

        self.btn_convert = ctk.CTkButton(self.top_frame, text="⚡ 변환 시작", command=self.start_conversion, state="disabled",
                                         fg_color="green", hover_color="darkgreen", height=35, width=120, font=("", 14, "bold"))
        self.btn_convert.pack(side="right", padx=5)

    def create_input_tree(self, parent):
        # 레이블
        label = ctk.CTkLabel(parent, text="📋 변환할 파일 목록", font=("", 13, "bold"), anchor="w")
        label.pack(fill="x", padx=5, pady=(5,0))

        # 프레임
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True)

        # 스크롤바 (가로/세로)
        ysb = ctk.CTkScrollbar(frame, orientation="vertical")
        ysb.pack(side="right", fill="y")
        
        xsb = ctk.CTkScrollbar(frame, orientation="horizontal")
        xsb.pack(side="bottom", fill="x")

        # Treeview
        # #0 컬럼을 체크박스+이미지용으로 사용 (show="tree headings")
        self.columns = ("name", "size", "mtime", "path")
        self.input_tree = ttk.Treeview(frame, columns=self.columns, show="tree headings", selectmode="extended",
                                       yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        
        ysb.configure(command=self.input_tree.yview)
        xsb.configure(command=self.input_tree.xview)
        
        self.input_tree.pack(side="left", fill="both", expand=True)

        # 헤더 설정
        self.input_tree.heading("#0", text="선택", command=self.toggle_all_checks)
        self.input_tree.heading("name", text="파일명", command=lambda: self.sort_tree(self.input_tree, "name", False))
        self.input_tree.heading("size", text="크기", command=lambda: self.sort_tree(self.input_tree, "size", False))
        self.input_tree.heading("mtime", text="수정일", command=lambda: self.sort_tree(self.input_tree, "mtime", False))
        self.input_tree.heading("path", text="폴더 위치", command=lambda: self.sort_tree(self.input_tree, "path", False))

        # 컬럼 너비
        self.input_tree.column("#0", width=50, anchor="center", stretch=False)
        self.input_tree.column("name", width=250, anchor="w")
        self.input_tree.column("size", width=80, anchor="center")
        self.input_tree.column("mtime", width=130, anchor="center")
        self.input_tree.column("path", width=300, anchor="w")

        # 이벤트
        self.input_tree.bind("<Button-1>", self.on_input_click) # 체크박스 토글
        self.input_tree.bind("<Delete>", lambda e: self.remove_checked_files())
        self.input_tree.bind("<Double-1>", lambda e: self.open_file(self.input_tree)) # 파일 열기

        # 우클릭 메뉴 (입력창용)
        self.input_menu = tk.Menu(self.input_tree, tearoff=0)
        self.input_menu.add_command(label="📄 파일 열기", command=lambda: self.open_file(self.input_tree))
        self.input_menu.add_command(label="📂 파일 위치 열기", command=lambda: self.open_folder(self.input_tree))
        self.input_menu.add_separator()
        self.input_menu.add_command(label="❌ 체크된 항목 제거", command=self.remove_checked_files)
        
        # 우클릭 바인딩 (헤더 판별 포함)
        self.input_tree.bind("<Button-3>", lambda e: self.on_tree_right_click(e, self.input_tree, self.input_menu))

    def create_result_tree(self, parent):
        label = ctk.CTkLabel(parent, text="✅ 변환 완료 목록", font=("", 13, "bold"), anchor="w", text_color="#4ade80")
        label.pack(fill="x", padx=5, pady=(5,0))

        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True)

        ysb = ctk.CTkScrollbar(frame, orientation="vertical")
        ysb.pack(side="right", fill="y")
        
        xsb = ctk.CTkScrollbar(frame, orientation="horizontal")
        xsb.pack(side="bottom", fill="x")

        # 결과창은 체크박스 불필요 (show="headings")
        cols = ("name", "size", "path")
        self.result_tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="extended",
                                        yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        
        ysb.configure(command=self.result_tree.yview)
        xsb.configure(command=self.result_tree.xview)
        
        self.result_tree.pack(side="left", fill="both", expand=True)

        self.result_tree.heading("name", text="PDF 파일명", command=lambda: self.sort_tree(self.result_tree, "name", False))
        self.result_tree.heading("size", text="크기", command=lambda: self.sort_tree(self.result_tree, "size", False))
        self.result_tree.heading("path", text="저장 위치", command=lambda: self.sort_tree(self.result_tree, "path", False))

        self.result_tree.column("name", width=250, anchor="w")
        self.result_tree.column("size", width=80, anchor="center")
        self.result_tree.column("path", width=400, anchor="w")

        # 이벤트
        self.result_tree.bind("<Double-1>", lambda e: self.open_file(self.result_tree)) # 파일 열기
        
        # 우클릭 메뉴 (결과창용)
        self.result_menu = tk.Menu(self.result_tree, tearoff=0)
        self.result_menu.add_command(label="📄 파일 열기", command=lambda: self.open_file(self.result_tree))
        self.result_menu.add_command(label="📂 파일 위치 열기", command=lambda: self.open_folder(self.result_tree))
        self.result_menu.add_separator()
        self.result_menu.add_command(label="✏️ 이름 변경", command=self.rename_result_file)
        self.result_menu.add_command(label="🗑️ 파일 삭제", command=self.delete_result_file)
        
        # 우클릭 바인딩
        self.result_tree.bind("<Button-3>", lambda e: self.on_tree_right_click(e, self.result_tree, self.result_menu))

    # --- 이벤트 핸들러 ---

    def on_tree_right_click(self, event, tree, body_menu):
        """우클릭 이벤트: 헤더인지 바디인지 구분하여 처리"""
        region = tree.identify_region(event.x, event.y)
        
        if region == "heading":
            # 헤더 우클릭 메뉴 표시
            col = tree.identify_column(event.x)
            self.current_header_col = col
            self.current_header_tree = tree
            self.header_menu.tk_popup(event.x_root, event.y_root)
        else:
            # 바디 우클릭 메뉴 표시
            item = tree.identify_row(event.y)
            if item:
                if item not in tree.selection():
                    tree.selection_set(item)
                body_menu.tk_popup(event.x_root, event.y_root)

    def autosize_current_col(self):
        if self.current_header_tree and self.current_header_col:
            self.autosize_column(self.current_header_tree, self.current_header_col)

    def autosize_all_cols(self):
        if self.current_header_tree:
            for col in self.current_header_tree["columns"]:
                self.autosize_column(self.current_header_tree, col)
            # 이름(#0) 컬럼이 있는 경우 (input_tree)
            if self.current_header_tree == self.input_tree:
                # 체크박스 컬럼은 고정
                pass

    def autosize_column(self, tree, col):
        """컬럼 너비 자동 조절 (폰트 측정 기반)"""
        # 체크박스 컬럼(#0)은 제외
        if col == "#0": return

        # 폰트 가져오기 (Treeview 스타일에서)
        try:
            style = ttk.Style()
            font_name = style.lookup("Treeview", "font")
            font = tk.font.Font(name=font_name) if font_name else tk.font.Font(family="", size=10) # 기본값
        except:
            font = tk.font.Font(family="", size=10)

        max_width = 0
        header_text = tree.heading(col)['text']
        # 헤더 텍스트 너비 고려 (헤더는 보통 bold이므로 약간 여유를 둠)
        max_width = font.measure(header_text) + 25 
        
        for item in tree.get_children():
            val = tree.set(item, col)
            # 실제 텍스트 너비 측정
            w = font.measure(str(val)) + 20 
            if w > max_width: max_width = w
        
        tree.column(col, width=max_width)

    def on_input_click(self, event):
        """입력 트리 클릭 (체크박스 토글 및 파일 열기 분기)"""
        region = self.input_tree.identify("region", event.x, event.y)
        if region == "tree": # #0 컬럼 영역 (체크박스)
            item = self.input_tree.identify_row(event.y)
            if item:
                tags = self.input_tree.item(item, "tags")
                if "checked" in tags:
                    self.input_tree.item(item, tags=("unchecked",), image=self.img_unchecked)
                else:
                    self.input_tree.item(item, tags=("checked",), image=self.img_checked)
        # 그 외 컬럼 클릭 시에는 기본 선택 동작(자동)

    def toggle_all_checks(self):
        """전체 체크 토글"""
        items = self.input_tree.get_children()
        if not items: return
        
        # 하나라도 언체크되어 있으면 모두 체크
        all_checked = True
        for item in items:
            if "unchecked" in self.input_tree.item(item, "tags"):
                all_checked = False
                break
        
        target_tag = "unchecked" if all_checked else "checked"
        target_img = self.img_unchecked if all_checked else self.img_checked
        
        for item in items:
            self.input_tree.item(item, tags=(target_tag,), image=target_img)

    # --- 파일 조작 기능 ---

    def open_file(self, tree):
        # 헤더 클릭 시 실행 방지 (더블클릭 이벤트가 겹칠 수 있음)
        if not tree.selection(): return
        
        sel = tree.selection()
        item = sel[0]
        vals = tree.item(item, "values")
        
        if tree == self.input_tree:
            path_idx, name_idx = 3, 0
        else:
            path_idx, name_idx = 2, 0
            
        full_path = os.path.join(vals[path_idx], vals[name_idx])
        try:
            os.startfile(full_path)
        except Exception as e:
            messagebox.showerror("오류", f"파일을 열 수 없습니다.\n{e}")

    def open_folder(self, tree):
        sel = tree.selection()
        if not sel: return
        item = sel[0]
        vals = tree.item(item, "values")
        
        if tree == self.input_tree:
            path_idx, name_idx = 3, 0
        else:
            path_idx, name_idx = 2, 0

        full_path = os.path.join(vals[path_idx], vals[name_idx])
        try:
            subprocess.run(['explorer', '/select,', os.path.abspath(full_path)])
        except Exception:
            pass

    def rename_result_file(self):
        sel = self.result_tree.selection()
        if not sel: return
        item = sel[0]
        vals = self.result_tree.item(item, "values") # name, size, path
        old_name = vals[0]
        folder = vals[2]
        old_path = os.path.join(folder, old_name)
        
        new_name = simpledialog.askstring("이름 변경", "새 파일 이름을 입력하세요:", initialvalue=old_name)
        if not new_name: return
        
        if not new_name.lower().endswith(".pdf"):
            new_name += ".pdf"
            
        new_path = os.path.join(folder, new_name)
        
        try:
            os.rename(old_path, new_path)
            # 트리 업데이트
            self.result_tree.set(item, "name", new_name)
        except Exception as e:
            messagebox.showerror("오류", f"이름 변경 실패: {e}")

    def delete_result_file(self):
        sel = self.result_tree.selection()
        if not sel: return
        
        if not messagebox.askyesno("삭제 확인", "선택한 파일을 정말 삭제하시겠습니까?\n복구할 수 없습니다."):
            return

        for item in sel:
            vals = self.result_tree.item(item, "values")
            path = os.path.join(vals[2], vals[0])
            try:
                os.remove(path)
                self.result_tree.delete(item)
            except Exception as e:
                print(f"삭제 실패: {e}")

    # --- 데이터 처리 ---

    def add_file_to_list(self, file_path):
        try:
            # 중복 체크
            for item in self.input_tree.get_children():
                vals = self.input_tree.item(item, "values")
                if os.path.abspath(os.path.join(vals[3], vals[0])) == os.path.abspath(file_path):
                    return

            stat = os.stat(file_path)
            size_str = self.format_size(stat.st_size)
            mtime_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(stat.st_mtime))
            
            # #0 컬럼에 이미지와 태그 설정
            self.input_tree.insert("", "end", text="", image=self.img_checked, tags=("checked",),
                                   values=(os.path.basename(file_path), size_str, mtime_str, os.path.dirname(file_path)))
            self.update_ui_states()
        except Exception:
            pass

    def add_result_item(self, pdf_path):
        """변환 성공한 파일을 결과창에 추가"""
        try:
            stat = os.stat(pdf_path)
            size_str = self.format_size(stat.st_size)
            self.result_tree.insert("", "end", values=(os.path.basename(pdf_path), size_str, os.path.dirname(pdf_path)))
        except:
            pass

    def format_size(self, size_bytes):
        if size_bytes < 1024: return f"{size_bytes} B"
        elif size_bytes < 1024**2: return f"{size_bytes/1024:.1f} KB"
        else: return f"{size_bytes/1024**2:.1f} MB"

    def parse_size(self, size_str):
        """파일 크기 문자열을 바이트 숫자로 변환"""
        try:
            parts = size_str.split()
            if len(parts) != 2: return 0
            num = float(parts[0])
            unit = parts[1].upper()
            if unit == "B": return num
            elif unit == "KB": return num * 1024
            elif unit == "MB": return num * 1024 * 1024
            elif unit == "GB": return num * 1024 * 1024 * 1024
            return 0
        except:
            return 0

    def remove_checked_files(self):
        items = [i for i in self.input_tree.get_children() if "checked" in self.input_tree.item(i, "tags")]
        for item in items:
            self.input_tree.delete(item)
        self.update_ui_states()

    def clear_file_list(self):
        for item in self.input_tree.get_children():
            self.input_tree.delete(item)
        self.update_ui_states()

    def sort_tree(self, tree, col, reverse):
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        
        if col == "size":
            # 숫자 기반 정렬
            l.sort(key=lambda t: self.parse_size(t[0]), reverse=reverse)
        else:
            # 기본 문자열 정렬
            l.sort(reverse=reverse)
            
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.sort_tree(tree, col, not reverse))

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("HWP 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")])
        for f in files: self.add_file_to_list(f)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith(('.hwp', '.hwpx')):
                        self.add_file_to_list(os.path.join(root, f))

    def process_command_line_args(self):
        if len(sys.argv) > 1:
            for arg in sys.argv[1:]:
                if os.path.exists(arg):
                    if os.path.isdir(arg):
                        for root, _, files in os.walk(arg):
                            for f in files:
                                if f.lower().endswith(('.hwp', '.hwpx')):
                                    self.add_file_to_list(os.path.join(root, f))
                    elif arg.lower().endswith(('.hwp', '.hwpx')):
                        self.add_file_to_list(arg)
            
            # 파일이 있으면 자동 시작
            if self.input_tree.get_children():
                self.after(500, self.start_conversion)

    def update_ui_states(self, is_converting=False):
        if not self.is_running: return
        try:
            state = "disabled" if is_converting else "normal"
            self.btn_files.configure(state=state)
            self.btn_folder.configure(state=state)
            self.btn_remove.configure(state=state)
            self.btn_clear.configure(state=state)
            
            if is_converting:
                self.btn_convert.configure(state="disabled")
            else:
                has_files = len(self.input_tree.get_children()) > 0
                self.btn_convert.configure(state="normal" if has_files else "disabled")
        except: pass

    def start_conversion(self):
        items = [i for i in self.input_tree.get_children() if "checked" in self.input_tree.item(i, "tags")]
        if not items:
            self.status_label.configure(text="⚠️ 선택된 파일이 없습니다.")
            return

        files = []
        for item in items:
            vals = self.input_tree.item(item, "values")
            files.append(os.path.join(vals[3], vals[0]))

        self.update_ui_states(True)
        threading.Thread(target=self.run_conversion, args=(files,), daemon=True).start()

    def run_conversion(self, files):
        from converter import convert_to_pdf
        total = len(files)
        success = 0
        
        for i, path in enumerate(files):
            if not self.is_running: return
            
            name = os.path.basename(path)
            self.after(0, lambda m=f"⏳ ({i+1}/{total}) 변환 중: {name}": self.status_label.configure(text=m))
            self.after(0, lambda v=(i+1)/total*0.9: self.progress_bar.set(v))
            
            # 결과 PDF 경로 예측
            pdf_path = os.path.splitext(path)[0] + ".pdf"
            
            if convert_to_pdf(path):
                success += 1
                # 결과창에 추가 (스레드 안전하게 after 사용)
                self.after(0, lambda p=pdf_path: self.add_result_item(p))
        
        if not self.is_running: return
        
        self.after(0, lambda: self.progress_bar.set(1.0))
        self.after(0, lambda: self.status_label.configure(text=f"✅ 변환 완료: {success}/{total} 성공"))
        self.after(0, lambda: self.update_ui_states(False))

    def destroy(self):
        self.is_running = False
        try: super().destroy()
        except: pass
        finally: sys.exit(0)

if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except: pass
    
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    
    app = App()
    app.mainloop()
