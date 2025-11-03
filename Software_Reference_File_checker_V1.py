import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Toplevel, Button
from bs4 import BeautifulSoup
import re
import time
import datetime
import pyautogui
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

       

class OVN_Execution(tk.Tk):

# fixed signature to accept base_path and proper super init
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

       
        self.title("Software Reference File checker")
        self.geometry("1360x400")
        self.base_path = base_path

        self.test_setups = ["HW1.6", "HW2.0", "HW2.5", "HW3.0", "HW3.0_Pre_C2.3"]

        self.Carlines = {
            "SP21": ["G60", "I20TP1"],
            "SP18": ["G26", "G08_LCI"],
            "SP21P": ["G70_LCI"]
        }

        self.selected_test_setup = tk.StringVar(self)
        self.selected_test_level = tk.StringVar(self)
        self.selected_test_module = tk.StringVar(self)
        self.selected_Carline = tk.StringVar(self)
        self.selected_test_type = tk.StringVar(self)
        self.selected_main_folder = tk.StringVar(self)
        self.selected_sub_folder = tk.StringVar(self)
        self.selected_sub_sub_folder = tk.StringVar(self)

        # Add trace to print value when changed
        self.selected_test_setup.trace_add('write', lambda *args: print(f"selected_HW: {self.selected_test_setup.get()}"))
        self.selected_Carline.trace_add('write', lambda *args: print(f"selected_Carline: {self.selected_Carline.get()}"))
        self.selected_main_folder.trace_add('write', lambda *args: print(f"selected_main_folder: {self.selected_main_folder.get()}"))
        self.selected_sub_folder.trace_add('write', lambda *args: print(f"selected_sub_folder: {self.selected_sub_folder.get()}"))
        self.selected_sub_sub_folder.trace_add('write', lambda *args: print(f"selected_sub_sub_folder: {self.selected_sub_sub_folder.get()}"))

        self.selected_main_folder.trace_add('write', lambda *args: self.varify_PICANDHWEL())
        self.selected_test_setup.trace_add('write', lambda *args: self.varify_PICANDHWEL())
        self.selected_sub_folder.trace_add('write', lambda *args: self.varify_PICANDHWEL())
        self.selected_sub_sub_folder.trace_add('write', lambda *args: self.varify_PICANDHWEL())
        self.selected_main_folder.trace_add('write', lambda *args: self.create_BMW_doc_path())
        self.selected_sub_folder.trace_add('write', lambda *args: self.create_BMW_doc_path())
        self.selected_sub_sub_folder.trace_add('write', lambda *args: self.create_BMW_doc_path())
        self.selected_main_folder.trace_add('write', lambda *args: self.checkref_candb_lindb_pdx())
        self.selected_sub_folder.trace_add('write', lambda *args: self.checkref_candb_lindb_pdx())
        self.selected_sub_sub_folder.trace_add('write', lambda *args: self.checkref_candb_lindb_pdx())
        self.selected_Carline.trace_add('write', lambda *args: self.checkref_swfk())


        

        self.setup_frame = ctk.CTkFrame(self, fg_color="transparent", border_width=1.5)
        self.setup_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)

        self.confirmation_frame = ctk.CTkFrame(self, fg_color="transparent", border_width=1.5)
        self.confirmation_frame.grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)

        self.main_folder_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_main_folder, values=[], height=23)
        self.sub_folder_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_sub_folder, values=[], height=23)
        self.sub_sub_folder_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_sub_sub_folder, values=[], height=23)
        self.test_module_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_test_module, values=[], height=23)
        self.carline_combobox = ctk.CTkOptionMenu(self.setup_frame, variable=self.selected_Carline, values=[], height=23)

        self.create_setup_widgets()
        self.create_confirmation_widgets()

        self.load_directory()
        self.icon_path = r"D:\SVN_ENG9\CANoe_Configuration\common\TestEnvironment_Eng9\HILS_TEST_AutoRun\log\Tools\CANoe_Icon.png"
        self.temp_filepath = None

    def create_setup_widgets(self):
        ttk.Label(self.setup_frame, text="Software Infomation", font=("Arial", 16, "bold")).grid(row=0, columnspan =2, pady=10)

        self.create_label_combobox_pair("Select HW:", self.test_setups, self.selected_test_setup, self.setup_frame, row=1, width=20)
        
        self.selected_main_folder.trace_add('write', lambda *args: self.update_sub_folders())
        self.selected_sub_folder.trace_add('write', lambda *args: self.update_sub_sub_folders())
        self.selected_test_level.trace_add('write', lambda *args: self.update_test_modules())
        self.selected_main_folder.trace_add('write', lambda *args: self.update_carlines())

        ttk.Label(self.setup_frame, text="Service Pack (SP):").grid(row=2, column=0, pady=5, padx=10, sticky=tk.W)
        ttk.Label(self.setup_frame, text="Start of Production (SOP):").grid(row=3, column=0, pady=5, padx=10, sticky=tk.W)
        ttk.Label(self.setup_frame, text="Software:").grid(row=4, column=0, pady=5, padx=10, sticky=tk.W)
        ttk.Label(self.setup_frame, text="Carline:").grid(row=5, column=0, pady=5, padx=10, sticky=tk.W)

        self.main_folder_combobox.grid(row=2, column=1, pady=5,)
        self.sub_folder_combobox.grid(row=3, column=1, pady=5)
        self.sub_sub_folder_combobox.grid(row=4, column=1, pady=5)
        self.carline_combobox.grid(row=5, column=1, pady=5)
    
    def create_confirmation_widgets(self):

        ttk.Label(self.confirmation_frame, text="Reference", font=("Arial", 16, "bold")).grid(row=0, columnspan=2, pady=10)

        # create labels that can be updated later (use globals safely with fallback)
        self.swfk_label = ttk.Label(self.confirmation_frame, text=f"⦿ SWFK: {SWFK or ''}")
        self.swfk_label.grid(row=3, column=0, pady=7, padx=10, sticky=tk.W)

        self.hwel_label = ttk.Label(self.confirmation_frame, text=f"⦿ HWEL: {HWEL or ''}")
        self.hwel_label.grid(row=4, column=0, pady=7, padx=10, sticky=tk.W)

        self.pic_label = ttk.Label(self.confirmation_frame, text=f"⦿ PIC: {PIC_Ver or ''}")
        self.pic_label.grid(row=5, column=0, pady=7, padx=10, sticky=tk.W)

        self.can_label = ttk.Label(self.confirmation_frame, text=f"⦿ CAN_NK: {CAN_NK or ''}")
        self.can_label.grid(row=6, column=0, pady=7, padx=10, sticky=tk.W)

        self.lin_label = ttk.Label(self.confirmation_frame, text=f"⦿ LIN_NK: {LIN_NK or ''}")
        self.lin_label.grid(row=7, column=0, pady=7, padx=10, sticky=tk.W)

        self.pdx_label = ttk.Label(self.confirmation_frame, text=f"⦿ PDX: {PDX or ''}")
        self.pdx_label.grid(row=8, column=0, pady=7, padx=10, sticky=tk.W)

#ใช้งานกับ create_setup_widgets
    def create_label_combobox_pair(self, label_text, items, variable, parent_frame, row, width):
        label = ttk.Label(parent_frame, text=label_text)
        label.grid(row=row, column=0, padx=10, pady=2, sticky=tk.W)
        combobox = ctk.CTkOptionMenu(parent_frame, values=items,variable=variable,height=23)
        combobox.grid(row=row, column=1, padx=10, pady=5)

#ใช้งานกับ create_confrimation_widgets
    def create_confirmation_frame(self, text, var, frame, row):

        ttk.Label(frame, text=text).grid(row=row, column=0, pady=7, padx=10, sticky=tk.W)

    def update_confirmation_labels(self):
        """Update displayed reference labels from globals."""
        try:
            self.swfk_label.config(text=f"⦿ SWFK: {SWFK or ''}")
            self.hwel_label.config(text=f"⦿ HWEL: {HWEL or ''}")
            self.pic_label.config(text=f"⦿ PIC: {PIC_Ver or ''}")
            self.can_label.config(text=f"⦿ CAN_NK: {CAN_NK or ''}")
            self.lin_label.config(text=f"⦿ LIN_NK: {LIN_NK or ''}")
            self.pdx_label.config(text=f"⦿ PDX: {PDX or ''}")
        except Exception:
            pass

#base functions for folder selection

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

            if sub_folders:
                self.selected_sub_folder.set(sub_folders[0])  # Set to the first sub folder by default
                self.update_sub_sub_folders()  # Update sub-sub folders for the initial selection

                # Update Carline options based on main folder selection
                self.update_carlines()

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

            if sub_sub_folders:
                self.selected_sub_sub_folder.set(sub_sub_folders[0])

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_test_modules(self):
        selected_test_level = self.selected_test_level.get()

        if selected_test_level in self.test_modules:
            test_module_options = self.test_modules[selected_test_level]
            self.selected_test_module.set("")  # Clear previous selection
            self.test_module_combobox.configure(values=test_module_options)

            if test_module_options:
                self.selected_test_module.set(test_module_options[0])  # Set to the first option by default

    def update_carlines(self):
        selected_main_folder = self.selected_main_folder.get()

        if selected_main_folder in self.Carlines:
            carline_options = self.Carlines[selected_main_folder]
            self.selected_Carline.set("")  # Clear previous selection
            self.carline_combobox.configure(values=carline_options)

            if carline_options:
                self.selected_Carline.set(carline_options[0])  # Set to the first option by default
        else:
            # If selected_main_folder doesn't match any keys in Carlines, clear the combobox values
            self.selected_Carline.set("")
            self.carline_combobox.configure(values=[])

    def load_directory(self):
        try:
            if not os.path.exists(self.base_path):
                messagebox.showerror("Error", f"Base path does not exist: {self.base_path}")
                return

            # List main folders
            self.main_folders = [name for name in os.listdir(self.base_path) if os.path.isdir(os.path.join(self.base_path, name))]
            self.main_folders.sort()
            self.selected_main_folder.set("")
            self.main_folder_combobox.configure(values=self.main_folders)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading directory: {str(e)}")


#Reference Check Functions
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

        for widget in self.confirmation_frame.grid_slaves():
            if isinstance(widget, ttk.Label) and "foreground" in widget.configure() and widget.cget("foreground") == "red":
                widget.destroy()

    def display_warning(self, message: str):
        
        self.clear_warning()
        warning_label = ttk.Label(self.confirmation_frame, text=message, foreground="red", font=("Arial", 10, "italic"))
        warning_label.grid(row=9, column=0, pady=5, padx=10, sticky=tk.W)

    def _norm(self, s: str) -> str:
        """Normalize version string from format like 'swfl_0000cded_000_006_001.msr' 
        to '0-6-1' format by extracting last 3 numbers."""
        if s is None:
            return ""
        # Convert to string, remove spaces, make lowercase
        s = os.path.basename(str(s)).replace(" ", "").strip().lower()
        
        # Find all number sequences in string
        numbers = re.findall(r"\d+", s)
        
        # Take last 3 numbers if available
        if len(numbers) >= 3:
            last_three = numbers[-3:]
            # Convert each number to int to remove leading zeros
            # Then back to string and join with '-'
            return "-".join(str(int(n)) for n in last_three)
        return ""

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

    def varify_PICANDHWEL(self) -> None:
        global HWEL, PIC_Ver, HW_Ver
        global pic_dir
        print("[DEBUG] varify_PICANDHWEL: Start")
        main = self.selected_main_folder.get() or ""
        sub = self.selected_sub_folder.get() or ""
        subsub = self.selected_sub_sub_folder.get() or ""
        print(f"[DEBUG] varify_PICANDHWEL: main={main}, sub={sub}, subsub={subsub}")

        # --- PIC ---

        try:
            pic_dir = os.path.join(
                self.base_path, main, sub, subsub,
                "Main SW", "BuildArtefact_Hotkey", "Esys", "BSW_MSR", "PIC"
            )
            print(f"[DEBUG] varify_PICANDHWEL: PIC dir={pic_dir}")
            files = []
            if os.path.isdir(pic_dir):
                # list only files (not directories) in pic_dir
                all_files = [f for f in os.listdir(pic_dir) if os.path.isfile(os.path.join(pic_dir, f))]
                if all_files:
                    # use the most recently modified file only
                    latest = max(all_files, key=lambda f: os.path.getmtime(os.path.join(pic_dir, f)))
                    files = [latest]
                else:
                    files = []
            print(f"[DEBUG] varify_PICANDHWEL: PIC files={files}")
            pic_ver_norm = self._norm(files[0]) if files else ""  # from test result (use first file if available)

            print(f"[DEBUG] varify_PICANDHWEL: PIC_Ver norm={pic_ver_norm}")
            PIC_Ver = pic_ver_norm  # Update global for consistency
            self.update_confirmation_labels()
            
        
                    
                    


        except Exception as e:
            print(f"[PIC verify] Exception: {e}")


        

        # --- HWEL ---
        # --- HWEL expectation logic ---
        expected_hwel_map = {
            "SP18":  {"HW1.6": "0000598F_005_001_007", "HW2.0": "0000598F_008_003_007", "HW2.5": "0000598F_008_005_007", "HW3.0": "-"},
            "SP21":  {"HW1.6": "0000598F_005_001_007", "HW2.0": "0000598F_008_003_007", "HW2.5": "0000598F_008_005_007", "HW3.0": "0000DDDC_010_006_007"},
            "SP21P": {"HW1.6": "-",                    "HW2.0": "0000BA75_108_003_007", "HW2.5": "0000BA75_008_005_007", "HW3.0": "0000DDF6_010_006_007"},
        }

        hw_key = self.selected_test_setup.get()
        print(f"[DEBUG] varify_PICANDHWEL: HW_Ver={self.selected_test_setup.get()}, hw_key={hw_key}")
        expect_hwel = None
        if main in expected_hwel_map and hw_key in expected_hwel_map[main]:
            expect_hwel = expected_hwel_map[main][hw_key]
        print(f"[DEBUG] varify_PICANDHWEL: expected HWEL={expect_hwel}")
        HWEL = expect_hwel  # Update global for consistency
        self.update_confirmation_labels()

    def checkref_candb_lindb_pdx(self) -> tuple[str, str, str]:
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
                    CAN_NK = expect_candb  # Update global for consistency
                    self.update_confirmation_labels()
                    break

        # Check LIN DB
        print(f"[DEBUG] check_candb_lindb_pdx: Checking NK path for LDF")
        if os.path.isdir(nk_path):
            for file in os.listdir(nk_path):
                if "LDF" in file:
                    expect_lindb = os.path.splitext(file)[0]
                    print(f"[DEBUG] check_candb_lindb_pdx: Found LIN DB={expect_lindb}")
                    LIN_NK = expect_lindb  # Update global for consistency
                    self.update_confirmation_labels()
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
                    PDX = expect_pdx  # Update global for consistency
                    self.update_confirmation_labels()
                    break
            
            # If no CCU-05 found, try DKC
            if not expect_pdx:
                for file in os.listdir(pdx_path):
                    if "DKC" in file and file.endswith(".pdx"):
                        expect_pdx = os.path.splitext(file)[0]
                        print(f"[DEBUG] check_candb_lindb_pdx: Found PDX (DKC)={expect_pdx}")
                        PDX = expect_pdx  # Update global for consistency
                        self.update_confirmation_labels()
                        break

       
        
        return expect_candb, expect_lindb, expect_pdx
    
    def checkref_swfk(self) -> tuple[str, bool]:
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
        SWFK = expect_swfk  # Update global SWFK to expected value for display
        print(f"new SWFK global {SWFK}")
        self.update_confirmation_labels()
        return expect_swfk



if __name__ == "__main__":

    file_path = r'D:\SVN_ENG9\CANoe_Configuration\common\TestEnvironment_Eng9\HILS_TEST_AutoRun\log'
    base_path = r"D:\SVN_ENG9\CCU_SW_Flash"
    app = OVN_Execution(base_path)
    app.mainloop()