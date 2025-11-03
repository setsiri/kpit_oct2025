import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


def _norm_str(x):
    if x is None:
        return None
    if isinstance(x, str):
        return x.strip()
    return str(x).strip()

def _cell_value(ws: Worksheet, coord: str):
    try:
        v = ws[coord].value
        return _norm_str(v)
    except Exception:
        return None

def Check_Yes_All_applicable_TC_Presented_And_No_Inapplicable_TC_Presented(ws: Worksheet, SP: str, SOP: str):
    """
    คืนค่าเป็น tuple: (Yes_All_applicable_TC_Presented, No_Inapplicable_TC_Presented)
    หรือ return 0 ถ้า row1/row2 หา matching ไม่เจอ
    """
    sheet_name = ws.title
    SP_norm = _norm_str(SP)
    SOP_norm = _norm_str(SOP)

    print(f"\n[DEBUG] ===== Sheet: {sheet_name} =====")
    print(f"[DEBUG] SP from path = {SP_norm}")
    print(f"[DEBUG] SOP from path = {SOP_norm}")

    # Row1: A1, J1, V1
    a1 = _cell_value(ws, 'A1')
    j1 = _cell_value(ws, 'J1')
    v1 = _cell_value(ws, 'V1')
    print(f"[DEBUG] Row1 values -> A1: {a1}, J1: {j1}, V1: {v1}")

    chosen_group = None  # 'A', 'J', 'V'
    if SP_norm and a1 and a1 == SP_norm:
        chosen_group = 'A'
    elif SP_norm and j1 and j1 == SOP_norm:  # ถ้าอยากเทียบ SP ใช้ j1 == SP_norm
        chosen_group = 'J'
    elif SP_norm and v1 and v1 == SP_norm:
        chosen_group = 'V'
    else:
        print("[DEBUG] row1 can't find matching")
        return 0

    # Row2 mapping
    if chosen_group == 'A':
        sop_map = [('D2', 'D'), ('E2', 'E'), ('F2', 'F'), ('G2', 'G')]
    elif chosen_group == 'J':
        sop_map = [('R2', 'R'), ('S2', 'S'), ('T2', 'T'), ('U2', 'U')]
    else:  # 'V'
        sop_map = [('V2', 'V')]

    chosen_col = None
    for coord, col in sop_map:
        val = _cell_value(ws, coord)
        print(f"[DEBUG] Checking {coord} = {val}")
        if SOP_norm and val and val == SOP_norm:
            chosen_col = col
            break

    if chosen_col is None:
        print("[DEBUG] row2 can't find matching")
        return 0

    print(f"[DEBUG] Matched SOP at column {chosen_col}")

    # Scan rows from row 3 downward
    max_r = ws.max_row or 3
    yes_rows = []
    yes_blank_ag = 0
    no_rows = []
    no_nonblank_ag = 0

    for r in range(3, max_r + 1):
        val = _cell_value(ws, f"{chosen_col}{r}")
        ag = _cell_value(ws, f"AG{r}")
        if val is None:
            continue
        vu = val.upper()
        if vu == "YES":
            yes_rows.append(r)
            if ag is None or ag == "":
                yes_blank_ag += 1
        elif vu == "NO":
            no_rows.append(r)
            if ag is not None and ag != "":
                no_nonblank_ag += 1

    yes_all = "Not all TC presented" if yes_rows and (yes_blank_ag > 0) else "All TC presented"
    no_inapp = "Inapplicable TC Presented" if no_rows and (no_nonblank_ag > 0) else "No Inapplicable TC Presented"

    print(f"[DEBUG] Total YES rows = {len(yes_rows)}, YES rows with AG blank = {yes_blank_ag}")
    print(f"[DEBUG] Total NO rows  = {len(no_rows)}, NO rows with AG non-blank = {no_nonblank_ag}")
    print(f"[DEBUG] => Yes_All_applicable_TC_Presented = {yes_all}")
    print(f"[DEBUG] => No_Inapplicable_TC_Presented    = {no_inapp}")

    return yes_all, no_inapp



def sw_info_tracking(path: str):
    """
    Normalize path แล้วดึงค่า SP, SOP, SW, CARLINE, HW, TEST_LEVEL ตามโครงสร้าง:
        .../Execution/{SP}/{SOP}/{SW}/{CARLINE}/{HW}/Tracking/{TEST_LEVEL}/...

    - รองรับทั้ง path ของไฟล์และโฟลเดอร์
    - ถ้าไม่ได้อยู่ตามโครงสร้างด้านบน จะพยายามหาแบบ heuristic จากชื่อโฟลเดอร์
    - พิมพ์ผลสำหรับ debug ใน terminal
    - คืนค่า dict สำหรับใช้งานต่อ (ถ้าต้องการ)
    """
    # 1) Normalize path
    norm = os.path.normpath(path)

    # ถ้าเป็นไฟล์ ให้ใช้โฟลเดอร์ที่ครอบมัน
    # ตรวจนามสกุลทั่วไปของ Excel หากตรงถือว่าเป็นไฟล์
    ext = os.path.splitext(norm)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        base_dir = os.path.dirname(norm)
    else:
        base_dir = norm  # เป็นโฟลเดอร์อยู่แล้ว หรือไม่ทราบแน่ชัดก็ถือเป็นโฟลเดอร์

    # แยกเป็นส่วนๆ ตามตัวคั่นของ OS
    parts = base_dir.split(os.sep)
    parts_lower = [p.lower() for p in parts]

    # 2) เตรียมตัวแปรผลลัพธ์
    SP = SOP = SW = CARLINE = HW = TEST_LEVEL = None

    # 3) พยายามอิงโครงแบบ canonical: .../Execution/SP/SOP/SW/CARLINE/HW/Tracking/TEST_LEVEL/...
    def idx_of_exact(name: str):
        try:
            return parts_lower.index(name.lower())
        except ValueError:
            return -1

    exec_idx = idx_of_exact("execution")
    tracking_idx = idx_of_exact("tracking")

    def safe_get(i):
        return parts[i] if 0 <= i < len(parts) else None

    if exec_idx != -1:
        candidate_SP = safe_get(exec_idx + 1)
        candidate_SOP = safe_get(exec_idx + 2)
        candidate_SW = safe_get(exec_idx + 3)
        candidate_CARLINE = safe_get(exec_idx + 4)
        candidate_HW = safe_get(exec_idx + 5)
        # ควรมี "Tracking" ที่ exec_idx + 6
        candidate_tracking = safe_get(exec_idx + 6)
        candidate_TEST_LEVEL = safe_get(exec_idx + 7)

        # Validate เบื้องต้นด้วย pattern
        if candidate_SP and candidate_SP.upper().startswith("SP"):
            SP = candidate_SP
        if candidate_SOP and candidate_SOP.upper().startswith("SOP"):
            SOP = candidate_SOP
        if candidate_SW:
            SW = candidate_SW
        if candidate_CARLINE:
            CARLINE = candidate_CARLINE
        if candidate_HW and candidate_HW.upper().startswith("HW"):
            HW = candidate_HW
        if (candidate_tracking and candidate_tracking.lower() == "tracking") and candidate_TEST_LEVEL:
            TEST_LEVEL = candidate_TEST_LEVEL

    # 4) ถ้าอย่างใดอย่างหนึ่งยังว่าง ลองหาแบบ heuristic (ค้นจากทุก segment)
    #    - SP: เริ่มด้วย SP ตามด้วยตัวเลข/ตัวอักษร
    #    - SOP: เริ่มด้วย SOP ตามด้วยตัวเลข/จุด
    #    - SW: ตัวเลขที่มีจุด เช่น 2535.0 (หรือเก็บตามชื่อโฟลเดอร์ถัดจาก SOP ถ้าเข้า pattern)
    #    - HW: เริ่มด้วย HW
    #    - TEST_LEVEL: โฟลเดอร์ถัดจาก "Tracking"
    if tracking_idx != -1 and TEST_LEVEL is None:
        TEST_LEVEL = safe_get(tracking_idx + 1)

    # หา SP
    if SP is None:
        for p in parts:
            if re.fullmatch(r"(?i)SP[\w.-]+", p):
                SP = p
                break

    # หา SOP
    if SOP is None:
        for i, p in enumerate(parts):
            if re.fullmatch(r"(?i)SOP[\w.-]+", p):
                SOP = p
                # เดาว่า SW อาจอยู่ถัดไปถ้ายังไม่มี
                if SW is None:
                    nxt = safe_get(i + 1)
                    if nxt:
                        SW = nxt
                break

    # หา HW
    if HW is None:
        for p in parts:
            if re.fullmatch(r"(?i)HW[\w.-]*", p):
                HW = p
                break

    # หา CARLINE (เดา: โฟลเดอร์ที่อยู่ระหว่าง SW กับ HW ถ้าเจอทั้งคู่)
    if CARLINE is None and SW and HW:
        try:
            i_sw = parts.index(SW)
            i_hw = parts.index(HW)
            if i_sw + 1 < i_hw:
                CARLINE = parts[i_sw + 1]
        except ValueError:
            pass

    # ถ้า SW ยังไม่เจอ ให้ลองหาโฟลเดอร์ที่เป็นรูปเลขมีจุด เช่น 2535.0
    if SW is None:
        for p in parts:
            if re.fullmatch(r"\d+(?:\.\d+)?", p):
                SW = p
                break

    # 5) Debug print (ตามที่ขอ ไม่โชว์บน GUI)
    print("\n[sw_info_tracking] ===========================")
    print(f"[sw_info_tracking] Normalized path : {norm}")
    print(f"[sw_info_tracking] Base directory  : {base_dir}")
    print(f"[sw_info_tracking] Parts            : {parts}")
    print(f"[sw_info_tracking] SP        = {SP}")
    print(f"[sw_info_tracking] SOP       = {SOP}")
    print(f"[sw_info_tracking] SW        = {SW}")
    print(f"[sw_info_tracking] CARLINE   = {CARLINE}")
    print(f"[sw_info_tracking] HW        = {HW}")
    print(f"[sw_info_tracking] TEST_LEVEL= {TEST_LEVEL}")
    print("[sw_info_tracking] ===========================\n")

    return {
        "SP": SP,
        "SOP": SOP,
        "SW": SW,
        "CARLINE": CARLINE,
        "HW": HW,
        "TEST_LEVEL": TEST_LEVEL,
        "normalized_path": norm,
        "base_dir": base_dir,
    }



def browse_file():
    path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xltx;*.xltm")],
    )
    if not path:
        return
    file_label.configure(text=path)
    fullpath_var.set(path)

    # เก็บค่าจาก path
    global info_vars
    info_vars = sw_info_tracking(path)

    # แสดงชื่อชีท + วิเคราะห์ผล
    show_sheets_and_analysis(path, info_vars)




def show_sheets_and_analysis(path: str, info: dict):
    textbox.configure(state="normal")
    textbox.delete("0.0", "end")
    try:
        wb = load_workbook(filename=path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        if not sheets:
            textbox.insert("end", "(No sheets found)\n")
        else:
            textbox.insert("end", "Sheets & Results:\n")
            textbox.insert("end", "----------------------------------------------\n")
            for i, s in enumerate(sheets, start=1):
                ws = wb[s]
                res = Check_Yes_All_applicable_TC_Presented_And_No_Inapplicable_TC_Presented(ws, info.get('SP'), info.get('SOP'))
                if res == 0:
                    textbox.insert("end", f"{i}. {s} -> row1/row2 can't find matching -> 0\n")
                else:
                    yes_all, no_inapp = res
                    textbox.insert("end", f"{i}. {s} -> Yes_All_applicable_TC_Presented = {yes_all}; No_Inapplicable_TC_Presented = {no_inapp}\n")
            textbox.insert("end", "----------------------------------------------\n")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read workbook:\n{e}")
    finally:
        textbox.configure(state="disabled")



def clear_selection():
    fullpath_var.set("")
    file_label.configure(text="No file selected")
    textbox.configure(state="normal")
    textbox.delete("0.0", "end")
    textbox.configure(state="disabled")


def main():
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Analysis Tracking Sheet")
    root.geometry("800x600")

    frame = ctk.CTkFrame(root, corner_radius=8)
    frame.pack(fill="both", expand=True, padx=16, pady=16)

    header = ctk.CTkLabel(frame, text="Analysis Tracking Sheet (Excel .xlsx)", font=ctk.CTkFont(size=16, weight="bold"))
    header.pack(pady=(8, 12))

    btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
    btn_frame.pack(fill="x", pady=(0, 8), padx=8)

    browse_btn = ctk.CTkButton(btn_frame, text="Browse...", width=120, command=browse_file)
    browse_btn.pack(side="left")

    clear_btn = ctk.CTkButton(btn_frame, text="Clear", width=80, command=clear_selection)
    clear_btn.pack(side="left", padx=(8, 0))

    global file_label, textbox, fullpath_var
    fullpath_var = tk.StringVar(value="")

    file_label = ctk.CTkLabel(frame, text="No file selected", anchor="w")
    file_label.pack(fill="x", padx=8, pady=(4, 8))

    # Textbox to show sheet names
    textbox = ctk.CTkTextbox(frame, width=520, height=200)
    textbox.pack(padx=8, pady=(0, 8), fill="both", expand=True)
    textbox.configure(state="disabled")

    root.mainloop()


if __name__ == "__main__":
    main()