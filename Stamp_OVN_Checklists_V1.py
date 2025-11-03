import os
import tkinter as tk
from tkinter import ttk, messagebox,filedialog,Toplevel, Button, messagebox
from bs4 import BeautifulSoup
import re
import time
import datetime
import pyautogui
import os
import subprocess
import pygetwindow as gw
import webbrowser
import tempfile
import base64
from io import BytesIO
import win32clipboard
import win32gui
import win32con
import customtkinter as ctk
import html
import sys

first_result = None
SWFK = None
HWEL = None
PIC_Ver = None
HW_Ver = None
CAN_NK = None
LIN_NK = None
PDX = None
Otherfailed = None
file_path = r'D:\SVN_ENG9\CANoe_Configuration\common\TestEnvironment_Eng9\HILS_TEST_AutoRun\log'
first_result_path = None
pic_dir = None #come form user input



#2
def load_1st_result(path):
    print(f"[DEBUG] load_1st_result: Start, path={path}")
    global SWFK,HWEL,PIC_Ver, HW_Ver, CAN_NK, LIN_NK, PDX, first_result_path
    SWFK = None
    HWEL = None
    PIC_Ver = None
    HW_Ver = None
    CAN_NK = None
    LIN_NK = None
    PDX = None
    first_result = None
    if not os.path.exists(file_path):
        first_folder = file_path
        print(f"[DEBUG] load_1st_result: file_path does not exist, using first_folder={first_folder}")
    else:
        for path in os.listdir(file_path):
            if os.path.isdir(os.path.join(file_path, path)) and path.startswith("ENG9"):
                first_folder = os.path.join(file_path, path)
                print(f"[DEBUG] load_1st_result: Found ENG9 folder, first_folder={first_folder}")
                break
    if not os.path.exists(first_folder):
        print(f"[DEBUG] load_1st_result: first_folder does not exist: {first_folder}")
        messagebox.showerror("Path not found", "Path of log folder is not found. Please select log folder manually.")
        sys.exit(1)
    else:
        for path in os.listdir(first_folder):
            print(f"[DEBUG] load_1st_result: file in folder: {path}")
            if os.path.isfile(os.path.join(first_folder, path)) and path.endswith(".html"):
                first_result = path
                print(f"[DEBUG] load_1st_result: Found first_result HTML file: {first_result}")
                break
    # If no test result html was found, show popup and exit
    if not first_result:
        messagebox.showerror("Test Result Not Found", "ไม่พบไฟล์ Test Result (.html) ในโฟลเดอร์ที่เลือก ระบบจะปิดโปรแกรมหลังจากกด OK")
        sys.exit(1)
    if first_result:
        first_result_path = os.path.join(first_folder, first_result)
        print(f"[DEBUG] load_1st_result: first_result_path={first_result_path}")
        with open(first_result_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            print(f"[DEBUG] load_1st_result: Read {len(lines)} lines from HTML file")
            Read_testlog(lines)


#3
def Read_testlog(lines):
    print(f"[DEBUG] Read_testlog: Start, lines={len(lines)}")
    global SWFK,HWEL,PIC_Ver, HW_Ver,CAN_NK,LIN_NK,PDX,Otherfailed
    OtherFailed = ""
    ErrInTestSystem_OtherFailed = ""
    line_num = 0
    hw_found = False  # Flag to track first HW: found
    for line in lines:
        line_num += 1
        # print(f"[DEBUG] Read_testlog: line {line_num}: {line.strip()}")
        if "SWFK [Logistics]" in line:
            SWFK = lines[line_num].replace('<td class="CellNoColor">','').replace('</td>','').strip()
            print(f"[DEBUG] Read_testlog: SWFK = {SWFK}")
        if "PIC version:" in line:
            PIC_Ver = lines[line_num].replace('<td class="CellNoColor">','').replace('</td>','').strip()
            print(f"[DEBUG] Read_testlog: PIC_Ver = {PIC_Ver}")
        
        if "HW:" in line and not hw_found:
            HW_Ver = html.unescape(lines[line_num]).replace('<td class="CellNoColor">', '').replace('</td>', '').strip()
            hw_found = True  # Set flag to True after first match
            print(f"[DEBUG] Read_testlog: HW_Ver = {HW_Ver}")


        if "HWEL [Logistics" in line:
            HWEL = lines[line_num].replace('<td class="CellNoColor">','').replace('</td>','').strip()
            print(f"[DEBUG] Read_testlog: HWEL = {HWEL}")
        if "Database AE_CAN_FD" in line:
            if  "BusDescription" in lines[line_num]:
                CAN_NK = re.sub(r'<td class="CellNoColor">.*?\\NK\\', '', lines[line_num])
                CAN_NK = re.sub(r'\(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\)', '', CAN_NK)
                CAN_NK = CAN_NK.replace('</td>', '').strip()
                print(f"[DEBUG] Read_testlog: CAN_NK = {CAN_NK}")
        if "Database LDF" in line:
            if  "LDF" in lines[line_num]:
                LIN_NK = re.sub(r'<td class="CellNoColor">.*?\\NK\\', '', lines[line_num])
                LIN_NK = re.sub(r'\(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\)', '', LIN_NK)
                LIN_NK = LIN_NK.replace('</td>', '').strip()
                print(f"[DEBUG] Read_testlog: LIN_NK = {LIN_NK}")
        if "Diagnostic Description of CCU_01:" in line:
            PDX = re.sub(r'<td class="CellNoColor">.*?\\PDX\\', '', lines[line_num])
            PDX = PDX.replace('</td>', '').strip()
            print(f"[DEBUG] Read_testlog: PDX = {PDX}")
        if ">Error in test system" in line:
            ErrInTestSystem_OtherFailed= lines[line_num].replace('<td class="NumberCell">','').replace('</td>','').strip()
            print(f"[DEBUG] Read_testlog: ErrInTestSystem_OtherFailed = {ErrInTestSystem_OtherFailed}")
        if '20px;">Fail' in line:
            OtherFailed= lines[line_num].replace('<td class="NumberCell">','').replace('</td>','').strip()
            print(f"[DEBUG] Read_testlog: OtherFailed = {OtherFailed}")
    if OtherFailed == "0" and ErrInTestSystem_OtherFailed == "0":
        Otherfailed = False
        print(f"[DEBUG] Read_testlog: Otherfailed = False")
    else:
        Otherfailed = True 
        print(f"[DEBUG] Read_testlog: Otherfailed = True")
    

    
    

class OVN_Execution(tk.Tk):
    def __init__(self, base_path):
        super().__init__()
        self.check_var_pic = tk.BooleanVar(self)
        self.check_var_hw = tk.BooleanVar(self)
        self.check_var_swfk = tk.BooleanVar(self)
        self.check_var_hwel = tk.BooleanVar(self)
        self.check_var_CAN = tk.BooleanVar(self)
        self.check_var_LIN = tk.BooleanVar(self)
        self.check_var_PDX = tk.BooleanVar(self)
        self.check_var_plain1 = tk.BooleanVar(self)
        self.check_var_plain2 = tk.BooleanVar(self)
        self.check_var_plain3 = tk.BooleanVar(self)
        self.check_var_plain4 = tk.BooleanVar(self)
        self.check_var_plain5 = tk.BooleanVar(self)
        self.check_var_plain6 = tk.BooleanVar(self)
        ctk.set_default_color_theme("green")
        self.title("Overnight Execution")
        self.geometry("1350x900")  
        self.base_path = base_path

        self.test_setups = ["TB1PUNE", "TB1BKK", "TB2", "TB3", "TB4", "TB5"]
        self.test_levels = ["FUSA","QM","ENG Release","HW2.0 Delta","FUSA+QM","HW1.6 Delta","HW3.0 Delta"]
        self.test_modules = {
            "FUSA": ["Charger", "14DCDC","CHG+DCDC", "NACS",],
            "QM": ["Charger", "14DCDC", "Smart Charging","CHG+DCDC"],
            "ENG Release" : ["Fault Injection"],
            "HW2.0 Delta" : ["ALL[Full]", "ALL[Failed]"],
            "HW1.6 Delta" : ["ALL[Full]", "ALL[Failed]"],
            "HW3.0 Delta" : ["ALL[Full]", "ALL[Failed]"],
            "FUSA+QM" : ["Fusa CHG+FUNC CHG", "Fusa 14DCDC + FUNC DCDC","Fusa CHG+FUNC DCDC","Fusa 14DCDC + FUNC CHG"]
        }
        self.Carlines = {
            "SP21": ["G60", "I20TP1"],
            "SP18": ["G26", "G08_LCI"],
            "SP21P": ["G70_LCI"]
        }
        self.test_types = ["Full Execution", "Failed Execution","Full+Failed Execution"]


        self.selected_test_setup = tk.StringVar(self)
        self.selected_test_level = tk.StringVar(self)
        self.selected_test_module = tk.StringVar(self)
        self.selected_Carline = tk.StringVar(self)
        self.selected_test_type = tk.StringVar(self)
        self.selected_main_folder = tk.StringVar(self)
        self.selected_sub_folder = tk.StringVar(self)
        self.selected_sub_sub_folder = tk.StringVar(self)

        # Add trace to print value when changed
        self.selected_test_setup.trace_add('write', lambda *args: print(f"selected_test_setup: {self.selected_test_setup.get()}"))
        self.selected_test_level.trace_add('write', lambda *args: print(f"selected_test_level: {self.selected_test_level.get()}"))
        self.selected_test_module.trace_add('write', lambda *args: print(f"selected_test_module: {self.selected_test_module.get()}"))
        self.selected_Carline.trace_add('write', lambda *args: print(f"selected_Carline: {self.selected_Carline.get()}"))
        self.selected_test_type.trace_add('write', lambda *args: print(f"selected_test_type: {self.selected_test_type.get()}"))
        self.selected_main_folder.trace_add('write', lambda *args: print(f"selected_main_folder: {self.selected_main_folder.get()}"))
        self.selected_sub_folder.trace_add('write', lambda *args: print(f"selected_sub_folder: {self.selected_sub_folder.get()}"))
        self.selected_sub_sub_folder.trace_add('write', lambda *args: print(f"selected_sub_sub_folder: {self.selected_sub_sub_folder.get()}"))

        self.selected_main_folder.trace_add('write', lambda *args: self.varify_PICANDHWEL())
        self.selected_sub_folder.trace_add('write', lambda *args: self.varify_PICANDHWEL())
        self.selected_sub_sub_folder.trace_add('write', lambda *args: self.varify_PICANDHWEL())
        self.selected_main_folder.trace_add('write', lambda *args: self.create_BMW_doc_path())
        self.selected_sub_folder.trace_add('write', lambda *args: self.create_BMW_doc_path())
        self.selected_sub_sub_folder.trace_add('write', lambda *args: self.create_BMW_doc_path())
        self.selected_main_folder.trace_add('write', lambda *args: self.check_candb_lindb_pdx())
        self.selected_sub_folder.trace_add('write', lambda *args: self.check_candb_lindb_pdx())
        self.selected_sub_sub_folder.trace_add('write', lambda *args: self.check_candb_lindb_pdx())
        self.selected_Carline.trace_add('write', lambda *args: self.check_swfk())

        self.setup_frame = ctk.CTkFrame(self,fg_color="transparent",border_width = 1.5)
        self.setup_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        
        self.selection_frame = ctk.CTkFrame(self,fg_color="transparent")
        self.selection_frame.grid(row=0, column=2, padx=5, pady=10, sticky=tk.W)
        
        
        self.confirmation_frame = ctk.CTkFrame(self,fg_color="transparent",border_width = 1.5)
        self.confirmation_frame.grid(row=1,column=0, columnspan =3, padx=10, pady=5, sticky=tk.W)
        
        self.screenshot_frame = ctk.CTkFrame(self,fg_color="transparent",border_width = 1.5)
        self.screenshot_frame.grid(row =0, column=1, padx=5, pady=5, sticky=tk.W)


        self.main_folder_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_main_folder, values=[],height=23,dropdown_fg_color = "#48E6A9",dropdown_hover_color="#48E6A9")
        self.sub_folder_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_sub_folder, values=[],height=23,dropdown_fg_color = "#48E6A9",dropdown_hover_color="#48E6A9")
        self.sub_sub_folder_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_sub_sub_folder, values=[],height=23,dropdown_fg_color = "#48E6A9",dropdown_hover_color="#48E6A9")
        self.test_module_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_test_module, values=[],height=23,dropdown_fg_color = "#48E6A9",dropdown_hover_color="#48E6A9")
        self.carline_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_Carline, values=[],height=23,dropdown_fg_color = "#48E6A9",dropdown_hover_color="#48E6A9")


        self.create_setup_widgets()
        self.create_confirmation_widgets()
        self.create_selection_widgets()
        #self.create_browse_widgets()
        self.create_screenshot_widgets()
        self.otherfailed()

        self.load_directory()
        self.icon_path = r"D:\SVN_ENG9\CANoe_Configuration\common\TestEnvironment_Eng9\HILS_TEST_AutoRun\log\Tools\CANoe_Icon.png"
        self.temp_filepath = None

    def create_setup_widgets(self):
        ttk.Label(self.setup_frame, text="OVN Information", font=("Arial", 14, "bold")).grid(row=0, columnspan =2, pady=10)
        self.create_label_combobox_pair("Select Test Setup:", self.test_setups, self.selected_test_setup, self.setup_frame, row=1, width=20)
        self.create_label_combobox_pair("Select Test Level:", self.test_levels, self.selected_test_level, self.setup_frame, row=6, width=20)
        self.create_label_combobox_pair("Select Test Type:", self.test_types, self.selected_test_type, self.setup_frame, row=8, width=20)


        self.selected_main_folder.trace_add('write', lambda *args: self.update_sub_folders())
        self.selected_sub_folder.trace_add('write', lambda *args: self.update_sub_sub_folders())
        self.selected_test_level.trace_add('write', lambda *args: self.update_test_modules())
        self.selected_main_folder.trace_add('write', lambda *args: self.update_carlines())

        ttk.Label(self.setup_frame, text="Service Pack (SP):").grid(row=2, column=0, pady=5, padx=10, sticky=tk.W)
        ttk.Label(self.setup_frame, text="Start of Production (SOP):").grid(row=3, column=0, pady=5, padx=10, sticky=tk.W)
        ttk.Label(self.setup_frame, text="Software:").grid(row=4, column=0, pady=5, padx=10, sticky=tk.W)
        ttk.Label(self.setup_frame, text="Test Module:").grid(row=7, column=0, pady=5, padx=10, sticky=tk.W)
        ttk.Label(self.setup_frame, text="Carline:").grid(row=5, column=0, pady=5, padx=10, sticky=tk.W)

        self.main_folder_combobox.grid(row=2, column=1, pady=5,)
        self.sub_folder_combobox.grid(row=3, column=1, pady=5)
        self.sub_sub_folder_combobox.grid(row=4, column=1, pady=5)
        self.test_module_combobox.grid(row=7, column=1, pady=5)
        self.carline_combobox.grid(row=5, column=1, pady=5)

    def create_browse_widgets(self):
        browse_labels = ["Select SWFK", "Select CAN_Database", "Select LIN_Database","Select PDX"]
        ttk.Label(self.browse_frame, text="Verify SW Information", font=("Arial", 14, "bold")).grid(row=0, columnspan=3, pady=10)
        for i in range(4):
            label_text = browse_labels[i]
            label = ttk.Label(self.browse_frame, text=f"{label_text}:")
            label.grid(row=9 + i, column=0, padx=10, pady=5,sticky=tk.W)
            
            entry = ctk.CTkEntry(self.browse_frame,width=225,fg_color="#37BD89",text_color ="white")
            entry.grid(row=9 + i, column=1, padx=10, pady=5)
            setattr(self, f"entry_{i+1}", entry)
            
           # button = ctk.CTkButton(self.browse_frame, text="Browse", command=lambda i=i: self.browse_file(i+1),width = 20)
           # button.grid(row=9 + i, column=2, padx=10, pady=5)
            
    def create_screenshot_widgets(self):

        ttk.Label(self.screenshot_frame, text="Attachment", font=("Arial", 14, "bold")).grid(row=0, columnspan =1, pady=10)
        self.txt_button = ctk.CTkButton(self.screenshot_frame, text="Take Screenshot of Init File", width=100, command=lambda: self.open_init_file("Init File"))
        self.txt_button.grid(row=1, column=0, padx=10, pady=5,sticky="nsew")

        self.html_button = ctk.CTkButton(self.screenshot_frame, text="Take Screenshot of 1st Testlog",width=100, command=lambda: self.open_html_file("1st Testlog"))
        self.html_button.grid(row=2, column=0, padx=10, pady=5,sticky="nsew")

        self.program_button = ctk.CTkButton(self.screenshot_frame, text="Take Screenshot of Test Unit",width=100, command=lambda: self.open_CANoe("Test Unit"))
        self.program_button.grid(row=3, column=0, padx=10, pady=5,sticky="nsew")
            
    


    def create_confirmation_widgets(self):
        try:
            ttk.Label(self.confirmation_frame, text="OVN Checklists", font=("Arial", 14, "bold")).grid(row=0, columnspan=2, pady=10)
            self.create_confirmation_frame("⦿ Clean up log folder(Y/N):", self.check_var_plain1, self.confirmation_frame, row=1)
            self.create_confirmation_frame("⦿ Get SelectGroup Folder, canoe.csv, init/init_first_execution from Original Folder (Y/N):", self.check_var_plain2, self.confirmation_frame, row=2)
            self.create_confirmation_frame("✪ SWFK is correct (Y/N): " + str(SWFK), self.check_var_swfk, self.confirmation_frame, row=3)
            self.create_confirmation_frame("⦿ HWEL is correct (Y/N): " + HWEL, self.check_var_hwel, self.confirmation_frame, row=4)
            self.create_confirmation_frame("⦿ PIC Version is correct (Y/N): "  + PIC_Ver, self.check_var_pic, self.confirmation_frame, row=5)
            self.create_confirmation_frame("⦿ CAN_NK selected in configuration is correct (Y/N): "+ CAN_NK, self.check_var_CAN, self.confirmation_frame, row=6)
            self.create_confirmation_frame("⦿ LIN_NK selected in configuration is correct (Y/N): "+ LIN_NK, self.check_var_LIN, self.confirmation_frame, row=7)
            self.create_confirmation_frame("⦿ PDX File selected in configuration is correct (Y/N): "+ PDX, self.check_var_PDX, self.confirmation_frame, row=8)
            self.create_confirmation_frame("⦿ All pre-condition sequence () in ENG9_Start pass (Y/N):", self.check_var_plain3, self.confirmation_frame, row=9)
            self.create_confirmation_frame("⦿ Confirm HW Version: "+ HW_Ver, self.check_var_hw, self.confirmation_frame, row=10)
            self.create_confirmation_frame("⦿ Check related test units that are executed (Y/N):", self.check_var_plain4, self.confirmation_frame, row=11)
            self.create_confirmation_frame("⦿ First result from Autorun has no other fail (Y/N)", self.check_var_plain5, self.confirmation_frame, row=12)
            self.create_confirmation_frame("⦿ Other Failed TCs included if any(NA if Full Execution is setup)", self.check_var_plain6, self.confirmation_frame, row=13)
            #button = ctk.CTkButton(self.confirmation_frame, text="Browse", command= self.browse_1st_result,width =10,state ="disabled")
           # button.grid(row=0, column=1, padx=10, pady=5)
  
        except:
            ttk.Label(self.confirmation_frame, text="OVN Checklists", font=("Arial", 14, "bold")).grid(row=0, columnspan=2, pady=10)
            self.create_confirmation_frame("⦿ Clean up log folder (Y/N):", self.check_var_plain1, self.confirmation_frame, row=1)
            self.create_confirmation_frame("⦿ Get SelectGroup Folder, canoe.csv, init/init_first_execution from Original Folder (Y/N):", self.check_var_plain2, self.confirmation_frame, row=2)
            self.create_confirmation_frame("✪ SWFK is correct (Y/N): " + str(SWFK), self.check_var_swfk, self.confirmation_frame, row=3)
            self.create_confirmation_frame("⦿ HWEL is correct (Y/N): ", self.check_var_hwel, self.confirmation_frame, row=4)
            self.create_confirmation_frame("⦿ PIC Version is correct (Y/N): ", self.check_var_pic, self.confirmation_frame, row=5)
            self.create_confirmation_frame("⦿ CAN_NK selected in configuration is correct (Y/N): ", self.check_var_CAN, self.confirmation_frame, row=6)
            self.create_confirmation_frame("⦿ LIN_NK selected in configuration is correct (Y/N): ", self.check_var_LIN, self.confirmation_frame, row=7)
            self.create_confirmation_frame("⦿ PDX File selected in configuration is correct (Y/N): ", self.check_var_PDX, self.confirmation_frame, row=8)
            self.create_confirmation_frame("⦿ All pre-condition sequence () in ENG9_Start pass (Y/N):", self.check_var_plain3, self.confirmation_frame, row=9)
            self.create_confirmation_frame("⦿ Confirm HW Version: ", self.check_var_hw, self.confirmation_frame, row=10)
            self.create_confirmation_frame("⦿ Check related test units that are executed (Y/N):", self.check_var_plain4, self.confirmation_frame, row=11)
            self.create_confirmation_frame("⦿ First result from Autorun has no other fail (Y/N)", self.check_var_plain5, self.confirmation_frame, row=12)
            self.create_confirmation_frame("⦿ Other Failed TCs included if any(NA if Full Execution is setup)", self.check_var_plain6, self.confirmation_frame, row=13)
            #button = ctk.CTkButton(self.confirmation_frame, text="Browse", command= self.browse_1st_result,width =10)
            #button.grid(row=0, column=1, padx=10, pady=5)

    def create_selection_widgets(self):
        self.show_button = ctk.CTkButton(
            self.selection_frame,
            text="Confirm Selection",
            command=self.validate_and_show,
            height = 80,
        )
        self.show_button.grid(row=0, column=0, columnspan=2, pady=5, padx=5)


#1 add on เลือก test result แบบ manaul
    def browse_1st_result(self):
        print("[DEBUG] browse_1st_result: Start")
        global file_path
        file_path = filedialog.askdirectory(initialdir="/", title="Select log folder")
        print(f"[DEBUG] browse_1st_result: file_path selected = {file_path}")
        if file_path:
            print("[DEBUG] browse_1st_result: Calling load_1st_result")
            load_1st_result(file_path)
            print("[DEBUG] browse_1st_result: Calling create_confirmation_widgets")
            self.create_confirmation_widgets()
        else:
            print("[DEBUG] browse_1st_result: No folder selected or dialog canceled.")
            messagebox.showerror("Error", f"No folder selected or dialog canceled.")
            



#ใช้งานกับ setup widgets
    def create_label_combobox_pair(self, label_text, items, variable, parent_frame, row, width):
        label = ttk.Label(parent_frame, text=label_text)
        label.grid(row=row, column=0, padx=10, pady=2, sticky=tk.W)
        combobox = ctk.CTkOptionMenu(parent_frame, values=items,variable=variable,height=23,dropdown_fg_color = "#2CB37F",dropdown_hover_color="#28B680")
        combobox.grid(row=row, column=1, padx=10, pady=5)


#ใช้งานกับ confrimation widgets
    def create_confirmation_frame(self, text, var, frame, row):
        ttk.Label(frame,text=text).grid(row=row, column=0, pady=5, padx=10, sticky=tk.W)
        ttk.Checkbutton(frame, text="Yes", variable=var).grid(row=row, column=1, pady=5, padx=10, sticky=tk.W)
        

#ใช้งานกับ confrimation widgets
    def otherfailed (self):
        global Otherfailed
        if Otherfailed is True:
            ttk.Checkbutton(self.confirmation_frame, text="No",state ='disabled').grid(row=11, column=1, pady=5, padx=10, sticky=tk.W)
            ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")).grid(row=11, column=1, pady=5, padx=10, sticky=tk.W)  
        else:
            self.check_var_plain5.set(True)
            ttk.Checkbutton(self.confirmation_frame, text="Yes", variable= self.check_var_plain5,state='disabled').grid(row=12, column=1, pady=5, padx=10, sticky=tk.W) 
  
  
  #ใช้กับ browe widget      
    def load_directory(self):
        try:
            # List main folders
            self.main_folders = [name for name in os.listdir(self.base_path) if os.path.isdir(os.path.join(self.base_path, name))]

            # Set initial options for main folder Combobox
            self.main_folders.sort()  # Sort main folders alphabetically
            self.selected_main_folder.set("")  # Clear previous selection
            self.main_folder_combobox.configure(values=self.main_folders)
            # ลบบรรทัดนี้เพื่อไม่ให้เลือกค่าแรกอัตโนมัติ
            # if self.main_folders:
            #     self.selected_main_folder.set(self.main_folders[0])
            #     self.update_sub_folders()
        except Exception as e:
                messagebox.showerror("Error", f"Error loading directory: {str(e)}")






    def update_sub_folders(self):
        # Update sub folder Combobox based on selected main folder
        selected_main_folder = self.selected_main_folder.get()
        main_folder_path = os.path.join(self.base_path, selected_main_folder)

        try:
            # List subfolders for the selected main folder
            sub_folders = [name for name in os.listdir(main_folder_path) if os.path.isdir(os.path.join(main_folder_path, name))]
            sub_folders.sort()  # Sort sub folders alphabetically
            self.selected_sub_folder.set("")  # Clear previous selection
            self.sub_folder_combobox.configure(values=sub_folders)
            # ลบบรรทัดนี้เพื่อไม่ให้เลือกค่าแรกอัตโนมัติ
            # if sub_folders:
            #     self.selected_sub_folder.set(sub_folders[0])
            #     self.update_sub_sub_folders()

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_sub_sub_folders(self):
        # Update sub-sub folder Combobox based on selected sub folder
        selected_main_folder = self.selected_main_folder.get()
        selected_sub_folder = self.selected_sub_folder.get()
        sub_folder_path = os.path.join(self.base_path, selected_main_folder, selected_sub_folder)

        try:
            # List sub-subfolders for the selected sub folder
            sub_sub_folders = [name for name in os.listdir(sub_folder_path) if os.path.isdir(os.path.join(sub_folder_path, name))]
            sub_sub_folders.sort()  # Sort sub-sub folders alphabetically
            self.selected_sub_sub_folder.set("")  # Clear previous selection
            self.sub_sub_folder_combobox.configure(values=sub_sub_folders)
            # ลบบรรทัดนี้เพื่อไม่ให้เลือกค่าแรกอัตโนมัติ
            # if sub_sub_folders:
            #     self.selected_sub_sub_folder.set(sub_sub_folders[0])

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_test_modules(self):
        selected_test_level = self.selected_test_level.get()

        if selected_test_level in self.test_modules:
            test_module_options = self.test_modules[selected_test_level]
            self.selected_test_module.set("")  # Clear previous selection
            self.test_module_combobox.configure(values=test_module_options)
            # ลบบรรทัดนี้เพื่อไม่ให้เลือกค่าแรกอัตโนมัติ
            # if test_module_options:
            #     self.selected_test_module.set(test_module_options[0])

    def update_carlines(self):
        selected_main_folder = self.selected_main_folder.get()

        if selected_main_folder in self.Carlines:
            carline_options = self.Carlines[selected_main_folder]
            self.selected_Carline.set("")  # Clear previous selection
            self.carline_combobox.configure(values=carline_options)
            # ลบบรรทัดนี้เพื่อไม่ให้เลือกค่าแรกอัตโนมัติ
            # if carline_options:
            #     self.selected_Carline.set(carline_options[0])
        else:
            # If selected_main_folder doesn't match any keys in Carlines, clear the combobox values
            self.selected_Carline.set("")  # Clear selection
            self.carline_combobox['values'] = []  # No values available





#ใช้กับ browe widget   และ confrimation widget
    def browse_file(self, index):
        global SWFK,CAN_NK,LIN_NK,PDX
        read_report_info = {
            1: SWFK,
            2: CAN_NK,
            3: LIN_NK,
            4: PDX,
        }
        browse_labels = {
            1: "Selected SWFK",
            2: "Selected CAN_Database",
            3: "Selected LIN_Database",
            4: "Selected PDX",
        }
        if index == 1:
            initial_dir = os.path.join(self.base_path, self.selected_main_folder.get(), self.selected_sub_folder.get(), self.selected_sub_sub_folder.get(),"CRETA")
            filetypes=[("S19 files", "*.s19")]
            
        elif index == 2:
            initial_dir = os.path.join(self.base_path, self.selected_main_folder.get(), self.selected_sub_folder.get(),"BMW_DOC")
            filetypes=[("arxml files", "*.arxml")]
            
        elif index == 3:
            initial_dir = os.path.join(self.base_path, self.selected_main_folder.get(), self.selected_sub_folder.get(),"BMW_DOC")
            filetypes=[("ldf files", "*.ldf")]
            
        elif index == 4:
            initial_dir = os.path.join(self.base_path, self.selected_main_folder.get(), self.selected_sub_folder.get(),"BMW_DOC")
            filetypes=[("pdx", "*.pdx")]
            
        entry = getattr(self, f"entry_{index}")
        file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=filetypes)
        displayed_text = os.path.basename(file_path)
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, displayed_text)
            
            if  read_report_info[index] not in displayed_text:
                messagebox.showerror("Incorrct SW Information", f"{browse_labels[index]} does not match with data: {read_report_info[index]} read from the test log")
                if index == 1:
                    ttk.Checkbutton(self.confirmation_frame, text="No",state ='disabled').grid(row=3, column=1, pady=5, padx=10, sticky=tk.W)
                    ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")).grid(row=3, column=1, pady=5, padx=10, sticky=tk.W)
                elif index == 2:
                    ttk.Checkbutton(self.confirmation_frame, text="No",state ='disabled').grid(row=6, column=1, pady=5, padx=10, sticky=tk.W)
                    ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")).grid(row=6, column=1, pady=5, padx=10, sticky=tk.W)
                elif index == 3:
                    ttk.Checkbutton(self.confirmation_frame, text="No",state ='disabled').grid(row=7, column=1, pady=5, padx=10, sticky=tk.W)
                    ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")).grid(row=7, column=1, pady=5, padx=10, sticky=tk.W)
                elif index == 4:
                    ttk.Checkbutton(self.confirmation_frame, text="No",state ='disabled').grid(row=8, column=1, pady=5, padx=10, sticky=tk.W)
                    ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")).grid(row=8, column=1, pady=5, padx=10, sticky=tk.W)
            else:
                if index == 1:
                    self.check_var_swfk.set(True)
                    ttk.Checkbutton(self.confirmation_frame, text="Yes", variable= self.check_var_swfk,state='disabled').grid(row=3, column=1, pady=5, padx=10, sticky=tk.W)
                elif index == 2:
                    self.check_var_CAN.set(True)
                    ttk.Checkbutton(self.confirmation_frame, text="Yes", variable= self.check_var_CAN,state='disabled').grid(row=6, column=1, pady=5, padx=10, sticky=tk.W)
                elif index == 3:
                    self.check_var_LIN.set(True)
                    ttk.Checkbutton(self.confirmation_frame, text="Yes", variable= self.check_var_LIN,state='disabled').grid(row=7, column=1, pady=5, padx=10, sticky=tk.W)
                elif index == 4:
                    self.check_var_PDX.set(True)
                    ttk.Checkbutton(self.confirmation_frame, text="Yes", variable= self.check_var_PDX,state='disabled').grid(row=8, column=1, pady=5, padx=10, sticky=tk.W)
     





    def open_init_file(self, file_type):
        txt_file_path = r"D:\SVN_ENG9\CANoe_Configuration\common\TestEnvironment_Eng9\HILS_TEST_AutoRun\Data\init.csv"
        
        if not os.path.isfile(txt_file_path):
            messagebox.showwarning("File Not Found", "The specified file was not found. Please select a file manually.")
            txt_file_path = filedialog.askopenfilename(title="Select File", filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
        
        if txt_file_path:
            self.open_file_and_capture_screenshot(txt_file_path,file_type)
        else:
            messagebox.showinfo("No File Selected", "No file was selected. Operation cancelled.")

    def open_html_file(self, file_type):
        html_file_path = first_result_path
        self.open_file_and_capture_screenshot(html_file_path, file_type)
            
    def open_CANoe(self, file_type):
        CANoe_Config = ["SP18.cfg", "SP21.cfg", "SP21P.cfg"]
        def enum_windows_callback(hwnd, windows):
            window_title = win32gui.GetWindowText(hwnd).lower()
            for title in CANoe_Config:
                if title.lower() in window_title:
                    windows.append(hwnd)
                    break  

        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)

        if windows:
            hwnd = windows[0]
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(3)
            print(f"Window with title containing one of {CANoe_Config} brought to the front.")
            self.capture_screenshot(file_type)
        else:
            print(f"No window with titles containing {CANoe_Config} found.")

    def open_file_and_capture_screenshot(self, file_path, file_type):
        try:
            if os.path.exists(file_path):
                # Open the file using the appropriate method based on file extension
                if file_path.endswith('.csv'):
                    process = subprocess.Popen(['notepad.exe', file_path])
                    print(f"Opened text file: {file_path}")

                    # Wait for Notepad to open and maximize
                    for _ in range(10):
                        notepad_window = None
                        windows = gw.getWindowsWithTitle(" - Notepad")
                        for window in windows:
                            if window.title.endswith(" - Notepad"):
                                notepad_window = window
                                break
                        if notepad_window:
                            notepad_window.maximize()
                            print("Maximized Notepad window")
                            break
                        time.sleep(2)  # Wait for 1 second before trying again
                    time.sleep(3)
                    # Capture screenshot and save
                    self.capture_screenshot(file_type)

                    # Close Notepad after taking the screenshot
                    process.terminate()

                elif file_path.endswith('.html'):
                    webbrowser.open(file_path)
                    print(f"Opened HTML file: {file_path}")

                    # Wait for a few seconds for the browser to open
                    time.sleep(5)
                    #chrome_window = gw.getWindowsWithTitle("Google Chrome")[0]
                    #chrome_window.maximize()

                    # Capture screenshot and save
                    self.capture_screenshot(file_type)

                else:
                    print(f"Unsupported file type: {file_path}")
                    return

            else:
                print(f"File not found: {file_path}")

        except Exception as e:
            print(f"Error opening file and capturing screenshot: {e}")

    def capture_screenshot(self, file_type):
        try:
            # Capture screenshot
            screenshot = pyautogui.screenshot()

            # Save the screenshot temporarily (optional)
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            self.temp_filepath = os.path.join(tempfile.gettempdir(), f"screenshot_{timestamp}.png")
            screenshot.save(self.temp_filepath)
            print(f"Screenshot saved as {self.temp_filepath}")

            # Convert screenshot to BMP format in memory
            output = BytesIO()
            screenshot.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()

            # Copy image to clipboard using win32clipboard
            self.copy_to_clipboard(win32clipboard.CF_DIB, data)

            print("Screenshot copied to clipboard successfully!")

            # Delete temporary screenshot file
            if self.temp_filepath and os.path.exists(self.temp_filepath):
                os.remove(self.temp_filepath)
                print(f"Deleted temporary file: {self.temp_filepath}")
                self.temp_filepath = None  # Reset temporary file path

            self.attributes("-topmost", True)
            self.update()
            self.attributes("-topmost", False)
            messagebox.showinfo("Success", f"Screenshot of {file_type} copied to clipboard successfully!", parent=self)

        except Exception as e:
            print(f"Error capturing and copying screenshot to clipboard: {e}")
            self.attributes("-topmost", True)
            self.update()
            self.attributes("-topmost", False)
            messagebox.showerror("Error", f"Error capturing and copying screenshot to clipboard: {e}",parent=self)


    def copy_to_clipboard(self, clip_type, data):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(clip_type, data)
            win32clipboard.CloseClipboard()

        except Exception as e:
            print(f"Error copying to clipboard: {e}")   


    def debug_print_selection(self):
            print("Debug Print Selection:")
            print(f"selected_test_setup: {self.selected_test_setup.get()}")
            print(f"selected_test_level: {self.selected_test_level.get()}")
            print(f"selected_test_module: {self.selected_test_module.get()}")
            print(f"selected_Carline: {self.selected_Carline.get()}")
            print(f"selected_test_type: {self.selected_test_type.get()}")
            print(f"selected_main_folder: {self.selected_main_folder.get()}")
            print(f"selected_sub_folder: {self.selected_sub_folder.get()}")
            print(f"selected_sub_sub_folder: {self.selected_sub_sub_folder.get()}")


    def validate_and_show(self):
    # Remove entry validation and only check combobox selections
        if (self.selected_main_folder.get() == "" or
            self.selected_sub_folder.get() == "" or
            self.selected_sub_sub_folder.get() == "" or
            self.selected_test_setup.get() == "" or
            self.selected_Carline.get() == "" or
            self.selected_test_level.get() == "" or
            self.selected_test_module.get() == "" or
            self.selected_test_type.get() == ""):
            messagebox.showerror("Incomplete Selection", "Please select all the necessary details")
        else:
            self.show_selection()
    
            
    def show_selection(self):
        selection_window = tk.Toplevel(self)
        selection_window.title("OVN Execution Checklist Report")
        selection_window.geometry("1000x850")

        information_frame = ctk.CTkFrame(selection_window,fg_color="transparent",border_width = 1.5,border_color="#2cc985")
        information_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)

        main_folder = self.selected_main_folder.get()
        sub_folder = self.selected_sub_folder.get()
        sub_sub_folder = self.selected_sub_sub_folder.get()
        test_setup = self.selected_test_setup.get()
        Carline = self.selected_Carline.get()
        test_level = self.selected_test_level.get()
        test_module = self.selected_test_module.get()
        test_type = self.selected_test_type.get()
        
        selection_label = ttk.Label(information_frame, text= "OVN Information", font=("Arial", 12, "bold"))
        selection_label.grid(row=0, column=0, pady=10, padx=10, sticky=tk.W)
        
        selection_label = ttk.Label(information_frame, text=f"OVN Started on: {test_setup}")
        selection_label.grid(row=1, column=0,pady=5, padx=10, sticky=tk.W)
        
        selection_label = ttk.Label(information_frame, text=f"SW: {main_folder} {sub_folder} {sub_sub_folder}")
        selection_label.grid(row=2, column=0,pady=5, padx=10, sticky=tk.W)
        
        selection_label = ttk.Label(information_frame, text=f"Carline: {Carline}")
        selection_label.grid(row=3, column=0, pady=5, padx=10, sticky=tk.W)
        
        selection_label = ttk.Label(information_frame, text=f"Test Level: {test_level}")
        selection_label.grid(row=4, column=0, pady=5, padx=10, sticky=tk.W)
        
        selection_label = ttk.Label(information_frame, text=f"Test Module: {test_module}")
        selection_label.grid(row=5, column=0, pady=5, padx=10, sticky=tk.W)
        
        selection_label = ttk.Label(information_frame, text=f"Test Type: {test_type}")
        selection_label.grid(row=6, column=0, pady=5, padx=10, sticky=tk.W)

        checkbox_frame = ctk.CTkFrame(selection_window,fg_color="transparent",border_width = 1.5,border_color="#2cc985")
        checkbox_frame.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)

        check_var_pic = self.check_var_pic.get()
        check_var_hw = self.check_var_hw.get()
        check_var_swfk = self.check_var_swfk.get()
        check_var_hwel = self.check_var_hwel.get()
        check_var_CAN = self.check_var_CAN.get()
        check_var_LIN = self.check_var_LIN.get()
        check_var_PDX = self.check_var_PDX.get()
        check_var_plain1 = self.check_var_plain1.get()
        check_var_plain2 = self.check_var_plain2.get()
        check_var_plain3 = self.check_var_plain3.get()
        check_var_plain4 = self.check_var_plain4.get()
        check_var_plain5 = self.check_var_plain5.get()
        check_var_plain6 = self.check_var_plain6.get()
        
        checkbox_text = ""
        checkbox_text += f"⦿ Clean up log folder (Y/N):{'️✔' if check_var_plain1 else '❌'}\n\n"
        checkbox_text += f"⦿ Get SelectGroup Folder, canoe.csv, init/init_first_execution from Original Folder (Y/N):  {'✔' if check_var_plain2 else '❌'}\n\n"
        checkbox_text += f"⦿ SWFK is correct (Y/N) : {SWFK}  {'✔' if check_var_swfk else '❌'}\n\n"
        checkbox_text += f"⦿ HWEL is correct (Y/N) : {HWEL}  {'✔' if check_var_hwel else '❌'}\n\n"
        checkbox_text += f"⦿ PIC Version is correct (Y/N) : {PIC_Ver} {'✔' if check_var_pic else '❌'}\n\n"
        checkbox_text += f"⦿ CAN_NK selected in configuration is correct (Y/N) : {CAN_NK} {'✔' if check_var_CAN else '❌'}\n\n"
        checkbox_text += f"⦿ LIN_NK selected in configuration is correct (Y/N) : {LIN_NK} {'✔' if check_var_LIN else '❌'}\n\n"
        checkbox_text += f"⦿ PDX_NK selected in configuration is correct (Y/N) : {PDX} {'✔' if check_var_PDX else '❌'}\n\n"
        checkbox_text += f"⦿ All pre-condition sequence () in ENG9_Start pass (Y/N):  {'✔' if check_var_plain3 else '❌'}\n\n"
        checkbox_text += f"⦿ Confirm HW Version : {HW_Ver} {'✔' if check_var_hw else '❌'}\n\n"
        checkbox_text += f"⦿ Check related test units that are executed (Y/N):  {'✔' if check_var_plain4 else '❌'}\n\n"
        checkbox_text += f"⦿ First result from Autorun has no other fail (Y/N):  {'✔' if check_var_plain5 else '❌'}\n\n"
        checkbox_text += f"⦿ Other Failed TCs included if any(NA if Full Execution is setup):  {'✔' if check_var_plain6 else '❌'}\n\n"
      
        

        checkbox_label = ttk.Label(checkbox_frame, text="OVN Checklist", font=("Arial", 12, "bold"))
        checkbox_label.grid(row=0, column=0, pady=10, padx=10, sticky=tk.W)

        checkbox_detail_label = ttk.Label(checkbox_frame, text=checkbox_text, justify="left", anchor="w")
        checkbox_detail_label.grid(row=1, column=0, pady=5, padx=10, sticky=tk.W)


#from setsiri
# 
# D:\SVN_ENG9\CCU_SW_Flash\SP21\SOP2603\BMW_Doc\CCUBEV21_2525\NK\LDF_SP2021_25KW17_V151_H_LIN_4_V318.ldf
# 
#         
    def _normalize_subsub(self) -> str:
        """Normalize sub-sub folder value to get first 4 digits.
        Returns empty string if input is invalid."""
        try:
            val = self.selected_sub_sub_folder.get()
            if not val:
                return ""
            # Extract first 4 digits from the string
            return val.split('.')[0][:4]
        except:
            return ""
        
    def create_BMW_doc_path(self) -> str:
        """
        Build BMW_Doc path and find a folder that matches selected_sub_sub_folder.
        Priority:
        1) folder name contains the full subsub (e.g. '2525.0' in 'ccu_2525.0')
        2) folder name contains progressively shortened versions of subsub until 4 digits remain.
        3) If multiple matches, select the folder with the highest numeric value in the subset.
        Returns matched folder full path or empty string if not found.
        """
        main = self.selected_main_folder.get() or ""
        sub = self.selected_sub_folder.get() or ""
        subsub = self.selected_sub_sub_folder.get() or ""
        print(f"[DEBUG] create_BMW_doc_path: main={main}, sub={sub}, subsub={subsub}")

        bmw_doc_dir = os.path.join(self.base_path, main, sub, "BMW_Doc")
        print(f"[DEBUG] create_BMW_doc_path: BMW_Doc dir = {bmw_doc_dir}")

        # Clear any existing warning message
        self.clear_warning()

        if not os.path.isdir(bmw_doc_dir):
            print(f"[DEBUG] create_BMW_doc_path: BMW_Doc dir does not exist")
            self.display_warning("                                                                                                \n                                                                                                                                       ")
            return ""

        lowered_subsub = subsub.lower()
        original_subsub = lowered_subsub

        best_match = None
        best_suffix_value = -1

        # Check for exact match first
        for entry in os.listdir(bmw_doc_dir):
            entry_path = os.path.join(bmw_doc_dir, entry)
            if not os.path.isdir(entry_path):
                continue

            if lowered_subsub in entry.lower():
                print(f"[DEBUG] create_BMW_doc_path: Found exact match -> {entry}")
                self.clear_warning()  # Clear warning if exact match is found
                self.display_warning("                                                                                                \n                                                                                                                                     ")
                return entry_path  # Return immediately if exact match is found

            match = re.search(rf"{re.escape(lowered_subsub)}(\d+)", entry.lower())
            if match:
                suffix_value = int(match.group(1))  # Convert the suffix to an integer
                print(f"[DEBUG] create_BMW_doc_path: Found match with suffix -> {entry}, suffix={suffix_value}")

                if suffix_value > best_suffix_value:
                    best_suffix_value = suffix_value
                    best_match = entry_path
                    print(f"[DEBUG] create_BMW_doc_path: Updated best match -> {best_match}, best_suffix_value={best_suffix_value}")

        # If exact match found, return immediately
        if best_match:
            print(f"[DEBUG] create_BMW_doc_path: Exact match found -> {best_match}")
            self.clear_warning()
            self.display_warning(
                f"Reference BMW_Doc path used is \n{best_match}"
            )
            return best_match

        while len(lowered_subsub) > 4:
            for entry in os.listdir(bmw_doc_dir):
                entry_path = os.path.join(bmw_doc_dir, entry)
                if not os.path.isdir(entry_path):
                    continue

                # Check if the current `lowered_subsub` is a subset of the folder name
                if lowered_subsub in entry.lower():
                    # Extract the suffix immediately following `lowered_subsub`
                    match = re.search(rf"{re.escape(lowered_subsub)}(\d+)", entry.lower())
                    if match:
                        suffix_value = int(match.group(1))  # Convert the suffix to an integer
                        print(f"[DEBUG] create_BMW_doc_path: Found match -> {entry}, suffix={suffix_value}")

                        # Update the best match if this suffix is larger
                        if suffix_value > best_suffix_value:
                            best_suffix_value = suffix_value
                            best_match = entry_path
                            print(f"[DEBUG] create_BMW_doc_path: Updated best match -> {best_match}, best_suffix_value={best_suffix_value}")

            # If match found after reduction, stop further reduction
            if best_match:
                print(f"[DEBUG] create_BMW_doc_path: Match found after reduction -> {best_match}")
                self.display_warning(
                    f"Reference BMW_Doc path used is \n{best_match}"
                )
                return best_match

            # Shorten `lowered_subsub` by removing the last character
            lowered_subsub = lowered_subsub[:-1]
            print(f"[DEBUG] create_BMW_doc_path: Shortened subsub to {lowered_subsub}")

        # Final attempt with exactly 4 characters
        if len(lowered_subsub) == 4:
            for entry in os.listdir(bmw_doc_dir):
                entry_path = os.path.join(bmw_doc_dir, entry)
                if not os.path.isdir(entry_path):
                    continue

                if lowered_subsub in entry.lower():
                    match = re.search(rf"{re.escape(lowered_subsub)}(\d+)", entry.lower())
                    if match:
                        suffix_value = int(match.group(1))
                        print(f"[DEBUG] create_BMW_doc_path: Found 4-digit match -> {entry}, suffix={suffix_value}")

                        if suffix_value > best_suffix_value:
                            best_suffix_value = suffix_value
                            best_match = entry_path
                            print(f"[DEBUG] create_BMW_doc_path: Updated best match -> {best_match}, best_suffix_value={best_suffix_value}")

        if best_match:
            print(f"[DEBUG] create_BMW_doc_path: Best match -> {best_match}")
            if lowered_subsub != original_subsub:
                self.display_warning(
                    f"Reference BMW_Doc path used is \n{best_match}"
                )
            return best_match

        print("[DEBUG] create_BMW_doc_path: No matching BMW_Doc subfolder found")
        self.display_warning("No matching BMW_Doc subfolder found.")
        return ""

    def clear_warning(self):
        """
        Clear any existing warning message in the GUI.
        """
        for widget in self.setup_frame.grid_slaves():
            if isinstance(widget, ttk.Label) and "foreground" in widget.configure() and widget.cget("foreground") == "red":
                widget.destroy()

    def display_warning(self, message: str):
        """
        Display a warning message in the GUI.
        """
        self.clear_warning()
        warning_label = ttk.Label(self.setup_frame, text=message, foreground="red", font=("Arial", 10, "italic"))
        warning_label.grid(row=10, column=0, columnspan=2, pady=5, padx=10, sticky=tk.W)

    def _norm(self, s: str) -> str:
        if s is None:
            return ""
        return os.path.basename(str(s)).replace(" ", "").strip().lower()

    def _normalize_hwver(self, s: str) -> str:
            if not s:
                return ""
            ss = s.replace(" ", "")
            print(f"Iam here {ss}")
    
            if ss in ["SOP1.5_C2.3", "1.6"]:
                return "HW1.6"
            elif ss in ["SOP3.0", "SOP3.0_C2.3"]:
                return "HW3.0"
            elif ss == "SOP2.0":
                return "HW2.0"
    
            for token in ("HW1.6", "HW2.0", "HW3.0"):
                if token.replace(".", "") in ss.replace(".", ""):
                    return token
            return ""

    def _grid_replace_at(self, row: int, column: int):
        for child in self.confirmation_frame.grid_slaves(row=row, column=column):
            child.destroy()

    def varify_PICANDHWEL(self) -> None:
        global pic_dir
        print("[DEBUG] varify_PICANDHWEL: Start")
        main = self.selected_main_folder.get() or ""
        sub = self.selected_sub_folder.get() or ""
        subsub = self.selected_sub_sub_folder.get() or ""
        print(f"[DEBUG] varify_PICANDHWEL: main={main}, sub={sub}, subsub={subsub}")

        ROW_HWEL = 4
        ROW_PIC  = 5

        # --- PIC ---
        pic_match = False
        try:
            pic_dir = os.path.join(
                self.base_path, main,sub, subsub,
                "Main SW", "BuildArtefact_Hotkey", "Esys", "BSW_MSR", "PIC"
            )
            print(f"[DEBUG] varify_PICANDHWEL: PIC dir={pic_dir}")
            files = []
            if os.path.isdir(pic_dir):
                files = [f for f in os.listdir(pic_dir) if os.path.isfile(os.path.join(pic_dir, f))]
            print(f"[DEBUG] varify_PICANDHWEL: PIC files={files}")
            pic_ver_norm = self._norm(PIC_Ver) #from test result

            print(f"[DEBUG] varify_PICANDHWEL: PIC_Ver norm={pic_ver_norm}")
            
            

            if pic_ver_norm and files:
                for f in files:
                    f_full = self._norm(f)
                    f_stem = self._norm(os.path.splitext(f)[0])
                    print(f"[DEBUG] varify_PICANDHWEL: Checking file: {f}, f_full={f_full}, f_stem={f_stem}")

                    results = []
                    for x in [f, f_full, f_stem]:
                        pic_numb_list = re.findall(r"\d+", x)
                        last_digits = [num[-1] for num in pic_numb_list[-3:]]
                        results.append("-".join(last_digits))
                    print(f"[DEBUG] varify_PICANDHWEL: files after extract: {results}")
                    
                    if pic_ver_norm in results:
                        pic_match = True
                        print(f"[DEBUG] varify_PICANDHWEL: PIC match found: {f}")

        except Exception as e:
            print(f"[PIC verify] Exception: {e}")

        print(f"[DEBUG] varify_PICANDHWEL: PIC match result={pic_match}")
        self._grid_replace_at(ROW_PIC, 1)
        if pic_match:
            self.check_var_pic.set(True)
            ttk.Checkbutton(self.confirmation_frame, text="Yes", variable=self.check_var_pic, state='disabled'
            ).grid(row=ROW_PIC, column=1, pady=5, padx=10, sticky=tk.W)
        else:
            ttk.Checkbutton(self.confirmation_frame, text="No", state='disabled'
            ).grid(row=ROW_PIC, column=1, pady=5, padx=10, sticky=tk.W)
            ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")
            ).grid(row=ROW_PIC, column=1, pady=5, padx=10, sticky=tk.W)

        # --- HWEL ---
        # --- HWEL expectation logic ---
        expected_hwel_map = {
            "SP18":  {"HW1.6": "0000598F_005_001_007", "HW2.0": ["0000598F_008_003_007", "0000598F_008_005_007"], "HW3.0": "-"},
            "SP21":  {"HW1.6": "0000598F_005_001_007", "HW2.0": ["0000598F_008_003_007", "0000598F_008_005_007"], "HW3.0": "0000DDDC_010_006_007"},
            "SP21P": {"HW1.6": "-",                    "HW2.0": ["0000BA75_108_003_007", "0000BA75_008_005_007"], "HW3.0": "0000DDF6_010_006_007"},
        }

        hw_key = self._normalize_hwver(HW_Ver)
        print(f"[DEBUG] varify_PICANDHWEL: HW_Ver={HW_Ver}, hw_key={hw_key}")
        expected = None
        if main in expected_hwel_map and hw_key in expected_hwel_map[main]:
            expected = expected_hwel_map[main][hw_key]
        print(f"[DEBUG] varify_PICANDHWEL: expected HWEL={expected}")

        hwel_match = False
        try:
            if expected in (None, "-", ""):
                hwel_match = (HWEL is None) or (str(HWEL).strip() in ("", "-", "None"))
                print(f"[DEBUG] varify_PICANDHWEL: HWEL expected empty, HWEL={HWEL}, hwel_match={hwel_match}")
                
            else:
                exp_norm = self._norm(expected)
                got_norm = self._norm(HWEL)
                
                print(f"[DEBUG] varify_PICANDHWEL: exp_norm={exp_norm}, got_norm={got_norm}")
                if exp_norm and got_norm:
                    hwel_match = (exp_norm == got_norm) or (exp_norm in got_norm) or (got_norm in exp_norm)
                    print(f"[DEBUG] varify_PICANDHWEL: HWEL match result={hwel_match}")
        except Exception as e:
            print(f"[HWEL verify] Exception: {e}")

        print(f"[DEBUG] varify_PICANDHWEL: HWEL match result={hwel_match}")
        self._grid_replace_at(ROW_HWEL, 1)
        if hwel_match:
            self.check_var_hwel.set(True)
            ttk.Checkbutton(self.confirmation_frame, text="Yes", variable=self.check_var_hwel, state='disabled'
            ).grid(row=ROW_HWEL, column=1, pady=5, padx=10, sticky=tk.W)
        else:
            ttk.Checkbutton(self.confirmation_frame, text="No", state='disabled'
            ).grid(row=ROW_HWEL, column=1, pady=5, padx=10, sticky=tk.W)
            ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")
            ).grid(row=ROW_HWEL, column=1, pady=5, padx=10, sticky=tk.W)

    def check_candb_lindb_pdx(self) -> tuple[str, str, str]:
        """
        Check for expected CAN DB, LIN DB and PDX files in BMW doc path.
        Returns tuple of (expect_candb, expect_lindb, expect_pdx).
        Each value will be empty string if not found.
        Also sets self.match_candb / self.match_lindb / self.match_pdx booleans
        comparing the expected values with CAN_NK / LIN_NK / PDX read from test log.
        """
        global CAN_NK, LIN_NK, PDX
        # Get BMW doc path internally
        bmw_doc_path = self.create_BMW_doc_path()
        
        print(f"[DEBUG] check_candb_lindb_pdx: Start with path={bmw_doc_path}")
        
        ROW_CAN = 6
        ROW_LIN = 7
        ROW_PDX = 8

        expect_candb = ""
        expect_lindb = ""
        expect_pdx = ""

        # initialize matches
        self.match_candb = False
        self.match_lindb = False
        self.match_pdx = False

        if not bmw_doc_path or not os.path.isdir(bmw_doc_path):
            print(f"[DEBUG] check_candb_lindb_pdx: Invalid BMW doc path")
            print(f"[DEBUG] check_candb_lindb_pdx: CAN_NK(from log)={CAN_NK}, LIN_NK={LIN_NK}, PDX={PDX}")
            return expect_candb, expect_lindb, expect_pdx

        # Check CAN DB
        nk_path = os.path.join(bmw_doc_path, "NK")
        print(f"[DEBUG] check_candb_lindb_pdx: Checking NK path={nk_path}")
        if os.path.isdir(nk_path):
            for file in os.listdir(nk_path):
                if "BusDescription" in file:
                    expect_candb = os.path.splitext(file)[0]
                    print(f"[DEBUG] check_candb_lindb_pdx: Found CAN DB={expect_candb}")
                    break

        # Check LIN DB
        print(f"[DEBUG] check_candb_lindb_pdx: Checking NK path for LDF")
        if os.path.isdir(nk_path):
            for file in os.listdir(nk_path):
                if "LDF" in file:
                    expect_lindb = os.path.splitext(file)[0]
                    print(f"[DEBUG] check_candb_lindb_pdx: Found LIN DB={expect_lindb}")
                    break

        # Check PDX
        pdx_path = os.path.join(bmw_doc_path, "PDX")
        print(f"[DEBUG] check_candb_lindb_pdx: Checking PDX path={pdx_path}")
        if os.path.isdir(pdx_path):
            # First try CCU-05
            for file in os.listdir(pdx_path):
                if "CCU-05" in file and file.endswith(".pdx"):
                    expect_pdx = os.path.splitext(file)[0]
                    print(f"[DEBUG] check_candb_lindb_pdx: Found PDX (CCU-05)={expect_pdx}")
                    break
            
            # If no CCU-05 found, try DKC
            if not expect_pdx:
                for file in os.listdir(pdx_path):
                    if "DKC" in file and file.endswith(".pdx"):
                        expect_pdx = os.path.splitext(file)[0]
                        print(f"[DEBUG] check_candb_lindb_pdx: Found PDX (DKC)={expect_pdx}")
                        break

        # Normalize helper using existing _norm
        def _n(x):
            return self._norm(x)

        expect_candb_norm = _n(expect_candb)
        expect_lindb_norm = _n(expect_lindb)
        expect_pdx_norm = _n(expect_pdx)

        can_nk_norm = _n(CAN_NK)
        lin_nk_norm = _n(LIN_NK)
        pdx_norm = _n(PDX)

        # Matching logic: expect is subset of reported or vice versa
        if expect_candb_norm and can_nk_norm:
            self.match_candb = (expect_candb_norm in can_nk_norm) or (can_nk_norm in expect_candb_norm)
        else:
            self.match_candb = False

        if expect_lindb_norm and lin_nk_norm:
            self.match_lindb = (expect_lindb_norm in lin_nk_norm) or (lin_nk_norm in expect_lindb_norm)
        else:
            self.match_lindb = False

        if expect_pdx_norm and pdx_norm:
            self.match_pdx = (expect_pdx_norm in pdx_norm) or (pdx_norm in expect_pdx_norm)
        else:
            self.match_pdx = False

        print(f"[DEBUG] check_candb_lindb_pdx: Final results:")
        print(f"  Expect CAN DB = '{expect_candb}' (norm='{expect_candb_norm}'), CAN_NK(from log)='{CAN_NK}' (norm='{can_nk_norm}'), match_candb={self.match_candb}")
        print(f"  Expect LIN DB = '{expect_lindb}' (norm='{expect_lindb_norm}'), LIN_NK(from log)='{LIN_NK}' (norm='{lin_nk_norm}'), match_lindb={self.match_lindb}")
        print(f"  Expect PDX    = '{expect_pdx}' (norm='{expect_pdx_norm}'), PDX(from log)='{PDX}' (norm='{pdx_norm}'), match_pdx={self.match_pdx}")

        if self.match_candb:
            self.check_var_CAN.set(True)
            ttk.Checkbutton(self.confirmation_frame, text="Yes", variable=self.check_var_CAN, state='disabled'
            ).grid(row=ROW_CAN, column=1, pady=6, padx=10, sticky=tk.W)
        else:
            ttk.Checkbutton(self.confirmation_frame, text="No", state='disabled'
            ).grid(row=ROW_CAN, column=1, pady=6, padx=10, sticky=tk.W)
            ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")
            ).grid(row=ROW_CAN, column=1, pady=6, padx=10, sticky=tk.W)

        if self.match_lindb:
            self.check_var_LIN.set(True)
            ttk.Checkbutton(self.confirmation_frame, text="Yes", variable=self.check_var_LIN, state='disabled'
            ).grid(row=ROW_LIN, column=1, pady=7, padx=10, sticky=tk.W)
        else:
            ttk.Checkbutton(self.confirmation_frame, text="No", state='disabled'
            ).grid(row=ROW_LIN, column=1, pady=7, padx=10, sticky=tk.W)
            ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")
            ).grid(row=ROW_LIN, column=1, pady=7, padx=10, sticky=tk.W)

        if self.match_pdx:
            self.check_var_PDX.set(True)
            ttk.Checkbutton(self.confirmation_frame, text="Yes", variable=self.check_var_PDX, state='disabled'
            ).grid(row=ROW_PDX, column=1, pady=8, padx=10, sticky=tk.W)
        else:
            ttk.Checkbutton(self.confirmation_frame, text="No", state='disabled'
            ).grid(row=ROW_PDX, column=1, pady=8, padx=10, sticky=tk.W)
            ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")
            ).grid(row=ROW_PDX, column=1, pady=8, padx=10, sticky=tk.W)
        
        return expect_candb, expect_lindb, expect_pdx
    
    def check_swfk(self) -> tuple[str, bool]:
        """
        Check for expected SWFK based on folder structure and carline.
        Returns tuple (expect_swfk, match_swfk).
        """
        global SWFK
        # Fixed indentation for entire function body
        main = self.selected_main_folder.get() or ""
        sub = self.selected_sub_folder.get() or ""
        subsub = self.selected_sub_sub_folder.get() or "" 
        carline = self.selected_Carline.get() or ""
        
        print(f"[DEBUG] check_swfk: main={main}, sub={sub}, subsub={subsub}, carline={carline}")
        
        # Normalize carline
        carline_normalized = ""
        if "G08" in carline:
            carline_normalized = "G08"
        elif "G26" in carline:
            carline_normalized = "G26"
        elif "G28" in carline:
            carline_normalized = "G28" 
        elif "G60" in carline:
            carline_normalized = "G6x"
        elif "I20" in carline:
            carline_normalized = "I20"
        elif "G70" in carline:
            carline_normalized = "G70"
            
        print(f"[DEBUG] check_swfk: carline_normalized={carline_normalized}")
        
        # Build base path
        creta_path = os.path.join(self.base_path, main, sub, subsub, "CRETA")
        print(f"[DEBUG] check_swfk: CRETA path={creta_path}")
        
        if not os.path.isdir(creta_path):
            print(f"[DEBUG] check_swfk: CRETA path does not exist")
            # set attributes for callers
            self.match_swfk = False
            return "", False
            
        # Find PV folders
        pv_folders = [f for f in os.listdir(creta_path) 
                     if os.path.isdir(os.path.join(creta_path, f)) and "PV" in f]
        print(f"[DEBUG] check_swfk: Found PV folders={pv_folders}")
        
        if not pv_folders:
            print(f"[DEBUG] check_swfk: No PV folders found")
            self.match_swfk = False
            return "", False
            
        selected_pv = ""
        if len(pv_folders) == 1:
            selected_pv = pv_folders[0]
        else:
            # Create popup for PV selection
            pv_window = tk.Toplevel()
            pv_window.title("Select PV Folder")
            pv_window.geometry("300x200")
            
            def on_pv_select():
                nonlocal selected_pv
                selected_pv = listbox.get(listbox.curselection())
                pv_window.destroy()
                
            listbox = tk.Listbox(pv_window)
            listbox.pack(pady=10)
            for pv in pv_folders:
                listbox.insert(tk.END, pv)
                
            tk.Button(pv_window, text="Select", command=on_pv_select).pack(pady=5)
            self.wait_window(pv_window)
        
        if not selected_pv:
            print(f"[DEBUG] check_swfk: No PV folder selected")
            self.match_swfk = False
            return "", False
            
        print(f"[DEBUG] check_swfk: Selected PV folder={selected_pv}")
        
        # Look for carline folder
        pv_path = os.path.join(creta_path, selected_pv)
        expect_swfk = ""
        
        for folder in os.listdir(pv_path):
            folder_path = os.path.join(pv_path, folder)
            if not os.path.isdir(folder_path):
                continue
                
            if carline_normalized == "I20":
                if "I20" in folder and "TOP" in folder and "TP1" in folder:
                    print(f"[DEBUG] check_swfk: Found I20 TOP folder={folder}")
                    # Look for .s19 file
                    for file in os.listdir(folder_path):
                        if file.endswith(".s19"):
                            expect_swfk = os.path.splitext(file)[0]
                            print(f"[DEBUG] check_swfk: Found SWFK={expect_swfk}")
                            break
                    break
            else:
                if carline_normalized in folder:
                    print(f"[DEBUG] check_swfk: Found carline folder={folder}")
                    # Look for .s19 file
                    for file in os.listdir(folder_path):
                        if file.endswith(".s19"):
                            expect_swfk = os.path.splitext(file)[0]
                            print(f"[DEBUG] check_swfk: Found SWFK={expect_swfk}")
                            break
                    break

        # Compare expect_swfk with global SWFK
        def _normalize_swfk(s: str) -> str:
            if not s:
                return ""
            s = str(s).strip().lower()
            # remove extension if present
            s = os.path.splitext(s)[0]
            return s

        expect_norm = _normalize_swfk(expect_swfk)
        swfk_norm = _normalize_swfk(SWFK)

        match_swfk = False
        if expect_norm and swfk_norm:
            # match if expect is substring of SWFK (or equal)
            match_swfk = expect_norm in swfk_norm or swfk_norm in expect_norm
        else:
            match_swfk = False

        # store attribute for later use and print debug
        self.match_swfk = match_swfk
        print(f"[DEBUG] check_swfk: expect_swfk='{expect_swfk}', expect_norm='{expect_norm}'")
        print(f"[DEBUG] check_swfk: SWFK(raw)='{SWFK}', SWFK_norm='{swfk_norm}'")
        print(f"[DEBUG] check_swfk: match_swfk={match_swfk}")

        ROW_SWFK = 3

        if match_swfk:
            self.check_var_swfk.set(True)
            ttk.Checkbutton(self.confirmation_frame, text="Yes", variable=self.check_var_swfk, state='disabled'
            ).grid(row=ROW_SWFK, column=1, pady=3, padx=10, sticky=tk.W)
        else:
            ttk.Checkbutton(self.confirmation_frame, text="No", state='disabled'
            ).grid(row=ROW_SWFK, column=1, pady=3, padx=10, sticky=tk.W)
            ttk.Label(self.confirmation_frame, text="X", font=("Arial", 13, "bold")
            ).grid(row=ROW_SWFK, column=1, pady=3, padx=10, sticky=tk.W)

        return expect_swfk, match_swfk

    
if __name__ == "__main__":
    load_1st_result(file_path)
    
    base_path = r"D:\SVN_ENG9\CCU_SW_Flash"  
    app = OVN_Execution(base_path)
    # Debug button for printing selected variables
    #debug_button = ctk.CTkButton(app, text="Debug Print Selection", command=app.debug_print_selection, fg_color="#FFAA00")
    #debug_button.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
    app.mainloop()