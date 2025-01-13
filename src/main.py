import datetime
import json
import sys
import os
import tkinter as tk
from tkinter import messagebox, Frame
from tkinter import ttk
from ttkbootstrap import Style
from PIL import Image, ImageTk
import pyglet
from merge import merge

CURRENT_VERSION = "1.0.0"

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

icon_path = resource_path("4i_icon.ico")

class MainWindow(tk.Tk):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.settings_path = resource_path("settings.json")
        self.settings = {}
        self.create_default_settings_if_needed()
        self.iconbitmap(icon_path)
        self.result_text_template_pay = (
            "1. 내용: {year_month_pay} {occupation_type} {insurance_type} 납부\n"
            "2. 금액: {total_amount}원\n"
            "\n"
            "붙임  {year_month_pay} {insurance_type} 고지내역 1부.  끝."
        )
        self.result_text_template = (
            "1. 내용: {year_month} {occupation_type} {insurance_type} {selected_fee_type} 수납\n"
            "2. 금액: {total_amount}원\n"
            "\n"
            "붙임  {year_month} {insurance_type} 고지내역 1부.  끝."
        )

        self.title("4대보험 취합 자동화 프로그램")
        self.geometry("1100x800")

        self.min_w = 50
        self.max_w = 200
        self.cur_width = self.min_w
        self.expanded = False
        self.pinned = tk.BooleanVar()
        self.pinned.set(True)

        self.font_path = resource_path("font/NotoSansKR-Medium.ttf")
        pyglet.options['win32_gdi_font'] = True
        pyglet.font.add_file(self.font_path)
        self.custom_font_family = "Noto Sans KR Medium"

        self.style = Style(theme='flatly')
        self.style.configure('.', background='white', foreground='black')
        self.style.configure("TFrame", background="white")
        self.style.configure("Custom.TButton", font=(self.custom_font_family, 13))
        self.style.configure("Custom.TLabel", font=(self.custom_font_family, 20))

        self.main_layout = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.main_layout.pack(fill=tk.BOTH, expand=True)

        self.sidebar_frame = Frame(self.main_layout, width=self.min_w, height=self.winfo_height())
        self.sidebar_frame.pack_propagate(False)
        self.sidebar_frame.pack(side=tk.LEFT, fill=tk.Y)

        self.icons = {
            "홈": ImageTk.PhotoImage(Image.open(resource_path("icons/home_icon.png")).resize((24, 24))),
            "급여 및 보험 자료 병합": ImageTk.PhotoImage(Image.open(resource_path("icons/merge_icon.png")).resize((24, 24))),
            "수납 내용 자동 생성": ImageTk.PhotoImage(Image.open(resource_path("icons/autogen_icon.png")).resize((24, 24))),
            "반환 내용 자동 생성": ImageTk.PhotoImage(Image.open(resource_path("icons/autogen_pay_icon.png")).resize((24, 24))),
            "설정": ImageTk.PhotoImage(Image.open(resource_path("icons/settings_icon.png")).resize((24, 24))),
            "도움말": ImageTk.PhotoImage(Image.open(resource_path("icons/help_icon.png")).resize((24, 24))),
        }

        self.create_sidebar(self.sidebar_frame)
        self.content_area = self.create_content_area()
        self.content_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.settings = self.load_settings()

        self.sidebar_frame.bind("<Enter>", self.expand)
        self.sidebar_frame.bind("<Leave>", self.contract)
        self.show_welcome()
        self.expand()
        self.toggle_pin_sidebar()
        
    def create_default_settings_if_needed(self):
        current_date = datetime.datetime.now()
        current_month = f"{current_date.year}년 {current_date.month}월"
        if not os.path.exists(self.settings_path):
            default_settings = {
                "file_path": "",                    
                "autogen": {   
                    "year_month": current_month,
                    "result_text_template": (
                        "1. 내용: {year_month} {occupation_type} {insurance_type} {selected_fee_type} 수납\n"
                        "2. 금액: {total_amount}원\n"
                        "\n"
                        "붙임  {year_month} {insurance_type} 고지내역 1부.  끝."
                    ),
                    "staff_types": {
                        "공무원": {"occupation_row": 9 - 1, "institution_row": 10 - 1},
                        "기간제교원": {"occupation_row": 11 - 1, "institution_row": 12 - 1},
                        "기타급여1 및 공무직 대체": {"occupation_row": 20 - 1, "institution_row": 21 - 1},
                        "기타급여4": {"occupation_row": 13 - 1, "institution_row": 14 - 1}
                    }
                },
                "autogen_pay": {
                    "year_month_pay": current_month,
                    "result_text_template_pay": (
                        "1. 내용: {year_month_pay} {occupation_type} {insurance_type} 납부\n"
                        "2. 금액: {total_amount}원\n"
                        "\n"
                        "붙임  {year_month_pay} {insurance_type} 고지내역 1부.  끝."
                    ),
                    "staff_types": ["공무원", "공무직원 및 기간제교원"],
                    "occupation_row_mapping": {
                        "공무원": {"건강보험": 8},
                        "공무직원 및 기간제교원": {"건강보험": 9, "국민연금": 10, "고용보험": 11, "산재보험": 12}
                    }
                }
            }

            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(default_settings, f, ensure_ascii=False, indent=4)        
        

    def show_welcome(self):
        self.clear_content_area()
        welcome_page_widget = Welcome_Page(self.content_area)
        welcome_page_widget.pack(fill=tk.BOTH, expand=True)

    def create_sidebar(self, parent):
        self.buttons = [
            ("홈", self.show_welcome),
            ("급여 및 보험 자료 병합", self.show_fouri_merger),
            ("수납 내용 자동 생성", self.show_fouri_autogen),
            ("반환 내용 자동 생성", self.show_fouri_autogen_pay),
            ("설정", self.show_settings),
            ("도움말", self.show_help),
        ]

        self.button_widgets = []
        for text, command in self.buttons:
            icon = self.icons[text]
            button = ttk.Button(parent, text=text, image=icon, compound='left', command=command, style="Custom.TButton")
            button.image = icon
            button.pack(fill=tk.X, pady=10)
            self.button_widgets.append((button, text))

        self.pin_checkbox = ttk.Checkbutton(parent, text="사이드바 고정", variable=self.pinned, command=self.toggle_pin_sidebar)
        self.pin_checkbox.pack(fill=tk.X, pady=10)

    def toggle_pin_sidebar(self):
        self.cur_width = self.max_w if self.pinned.get() else self.min_w
        self.sidebar_frame.config(width=self.cur_width)
        self.fill_buttons(force_expand=self.pinned.get())

    def expand(self, event=None):
        if not self.pinned.get() and self.cur_width < self.max_w:
            self.cur_width += 15
            self.sidebar_frame.config(width=self.cur_width)
            self.after(3, self.expand)
        else:
            self.expanded = True
            self.fill_buttons(force_expand=self.pinned.get())

    def contract(self, event=None):
        if not self.pinned.get() and self.sidebar_frame.winfo_containing(*self.winfo_pointerxy()) == self.sidebar_frame:
            return
        if not self.pinned.get() and self.cur_width > self.min_w:
            self.cur_width -= 15
            self.sidebar_frame.config(width=self.cur_width)
            self.after(3, self.contract)
        else:
            self.expanded = False
            self.fill_buttons(force_expand=self.pinned.get())

    def fill_buttons(self, force_expand: bool = False) -> None:
        MIN_BUTTON_WIDTH = 2
        MAX_BUTTON_WIDTH = 20

        for button, text in self.button_widgets:
            width = MAX_BUTTON_WIDTH if self.expanded or force_expand else MIN_BUTTON_WIDTH
            button.config(text=text if self.expanded or force_expand else "", width=width)


    def create_content_area(self):
        return ttk.Frame(self.main_layout)

    def show_fouri_merger(self):
        self.geometry("1050x650")
        self.clear_content_area()
        fouri_merger_page = merge(self.content_area)
        fouri_merger_page.pack(fill=tk.BOTH, expand=True)

    def show_help(self):
        self.clear_content_area()
        help_page_widget = Help_Page(self.content_area)
        help_page_widget.pack(fill=tk.BOTH, expand=True)

    def clear_content_area(self):
        for widget in self.content_area.winfo_children():
            widget.destroy()

    def show_settings(self):
        self.geometry("1200x900")
        self.clear_content_area()

        self.settings_page = ttk.Frame(self.content_area)
        self.settings_page.pack(fill=tk.BOTH, expand=True)

        font = (self.custom_font_family, 14)

        file_path_label = ttk.Label(self.settings_page, text="파일 경로:", font=font)
        file_path_label.grid(row=0, column=0, sticky=tk.W, padx=20, pady=20)
        self.file_path_entry = ttk.Entry(self.settings_page, font=font)
        self.file_path_entry.grid(row=0, column=1, padx=20, pady=20, sticky=tk.EW)

        current_date = datetime.datetime.now()
        current_month = f"{current_date.year}년 {current_date.month}월"
        month_label = ttk.Label(self.settings_page, text="수납 연월:", font=font)
        month_label.grid(row=1, column=0, sticky=tk.W, padx=20, pady=20)
        self.month_entry = ttk.Entry(self.settings_page, font=font)
        self.month_entry.grid(row=1, column=1, padx=20, pady=20, sticky=tk.EW)

        pay_month_label = ttk.Label(self.settings_page, text="반환 연월:", font=font)
        pay_month_label.grid(row=2, column=0, sticky=tk.W, padx=20, pady=20)
        self.pay_month_entry = ttk.Entry(self.settings_page, font=font)
        self.pay_month_entry.grid(row=2, column=1, padx=20, pady=20, sticky=tk.EW)

        template_label = ttk.Label(self.settings_page, text="수납 내용:", font=font)
        template_label.grid(row=3, column=0, sticky=tk.W, padx=20, pady=20)
        self.template_text = tk.Text(self.settings_page, height=7, width=50, font=font, wrap=tk.WORD)
        self.template_text.grid(row=3, column=1, padx=20, pady=20, sticky=tk.EW)

        template_pay_label = ttk.Label(self.settings_page, text="반환 내용:", font=font)
        template_pay_label.grid(row=4, column=0, sticky=tk.W, padx=20, pady=20)
        self.template_text_pay = tk.Text(self.settings_page, height=7, width=50, font=font, wrap=tk.WORD)
        self.template_text_pay.grid(row=4, column=1, padx=20, pady=20, sticky=tk.EW)

        drag_drop_frame = ttk.Frame(self.settings_page)
        drag_drop_frame.grid(row=5, column=1, padx=20, pady=20)

        variables = {
            "수납연월": "{year_month}",
            "반환연월": "{year_month_pay}",
            "직종": "{occupation_type}",
            "보험 종류": "{insurance_type}",
            "개인 혹은 기관부담금": "{selected_fee_type}",
            "금액": "{total_amount}"
        }

        for display_name, variable in variables.items():
            label = ttk.Label(drag_drop_frame, text=display_name, background="lightgrey", width=11, anchor=tk.CENTER, font=font)
            label.pack(side=tk.LEFT, padx=5, pady=5)
            label.bind("<Button-1>", lambda e, var=variable: self.on_drag_start(e, var))

        button_frame = ttk.Frame(self.settings_page)
        button_frame.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky=tk.EW)

        save_button = ttk.Button(button_frame, text="저장", command=self.save_settings, bootstyle="success", style="Custom.TButton")
        save_button.pack(side=tk.LEFT, padx=5)
        load_button = ttk.Button(button_frame, text="불러오기", command=self.load_settings, bootstyle="info", style="Custom.TButton")
        load_button.pack(side=tk.LEFT, padx=5)

        self.load_settings()

    def on_drag_start(self, event, var):
        widget = self.content_area.focus_get()
        if widget in (self.template_text, self.template_text_pay):
            widget.insert(tk.INSERT, var)
        

    def save_settings(self):
        settings = {
            "file_path": self.file_path_entry.get(),    
            "autogen": {        
                "year_month": self.month_entry.get(),
                "result_text_template": self.template_text.get("1.0", tk.END).strip(),
                "staff_types": self.settings.get("autogen", {}).get("staff_types", [])
            },
            "autogen_pay": {
                "year_month_pay": self.pay_month_entry.get(),
                "result_text_template_pay": self.template_text_pay.get("1.0", tk.END).strip(),
                "staff_types": self.settings.get("autogen_pay", {}).get("staff_types", []),
                "occupation_row_mapping": self.settings.get("autogen_pay", {}).get("occupation_row_mapping", {})
            }
        }
        with open(self.settings_path, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)

        messagebox.showinfo('설정', '설정 저장 완료.')
        
    def update_settings(self, new_settings):
        self.settings.update(new_settings)

        if hasattr(self, 'settings_page') and self.settings_page:
            self.month_entry.delete(0, tk.END)
            self.month_entry.insert(0, self.settings["year_month"])        

    def load_settings(self, show_message=True):
        current_date = datetime.datetime.now()
        current_month = f"{current_date.year}년 {current_date.month}월"
        default_settings = {
            "file_path": "",              
            "autogen": {              
                "year_month":current_month,
                "result_text_template": (
                    "1. 내용: {year_month} {occupation_type} {insurance_type} {selected_fee_type} 수납\n"
                    "2. 금액: {total_amount}원\n"
                    "\n"
                    "붙임  {year_month} {insurance_type} 고지내역 1부.  끝."
                ),
                "staff_types": ["공무원", "공무직원 및 기간제교원"],
                "occupation_row_mapping": {
                    "공무원": {"건강보험": 8}
                }
            },
            "autogen_pay": {
                "year_month_pay": current_month,
                "result_text_template_pay": (
                    "1. 내용: {year_month_pay} {occupation_type} {insurance_type} 납부\n"
                    "2. 금액: {total_amount}원\n"
                    "\n"
                    "붙임  {year_month_pay} {insurance_type} 고지내역 1부.  끝."
                ),
                "staff_types": ["공무원", "공무직원 및 기간제교원"],
                "occupation_row_mapping": {
                    "공무원": {"건강보험": 8},
                    "공무직원 및 기간제교원": {"건강보험": 9, "국민연금": 10, "고용보험": 11, "산재보험": 12}
                }
            }
        }

        try:
            with open(self.settings_path, "r", encoding="utf-8") as f:
                loaded_settings = json.load(f)
                default_settings.update(loaded_settings)

            if hasattr(self, 'settings_page') and self.settings_page:
                self.file_path_entry.delete(0, tk.END)
                self.file_path_entry.insert(0, default_settings.get("file_path", ""))

                self.month_entry.delete(0, tk.END)
                self.month_entry.insert(0, default_settings.get("autogen", {}).get("year_month", ""))

                self.template_text.delete("1.0", tk.END)
                self.template_text.insert(tk.END, default_settings.get("autogen", {}).get("result_text_template", ""))

                self.pay_month_entry.delete(0, tk.END)
                self.pay_month_entry.insert(tk.END, default_settings.get("autogen_pay", {}).get("year_month_pay", ""))

                self.template_text_pay.delete("1.0", tk.END)
                self.template_text_pay.insert(tk.END, default_settings.get("autogen_pay", {}).get("result_text_template_pay", ""))

            self.settings.update(default_settings)

        except Exception as e:
            messagebox.showerror('설정', f'설정 파일을 불러오는 중 오류가 발생했습니다: {e}')

        return default_settings

    def show_fouri_autogen(self):
        self.clear_content_area()
        self.geometry("1100x800")
        fouri_autogen_page = fouri_autogen(
            self.content_area, 
            file_path=self.settings.get("file_path", ""),  
            year_month=self.settings.get("autogen", {}).get("year_month", ""),
            result_text_template=self.settings.get("result_text_template", ""),
        )
        fouri_autogen_page.pack(fill=tk.BOTH, expand=True)

    def show_fouri_autogen_pay(self):
        self.clear_content_area()
        self.geometry("1100x800")
        fouri_autogen_pay_page = fouri_autogen_pay(
            self.content_area, 
            file_path=self.settings.get("file_path", ""),  
            year_month_pay=self.settings.get("autogen_pay", {}).get("year_month_pay", ""),
            result_text_template_pay=self.settings.get("autogen_pay", {}).get("result_text_template_pay", ""),
            staff_types=self.settings.get("autogen_pay", {}).get("staff_types", []),
            occupation_row_mapping=self.settings.get("autogen_pay", {}).get("occupation_row_mapping", {})
        )
        fouri_autogen_pay_page.pack(fill=tk.BOTH, expand=True)

if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
