import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import logging
import serial.tools.list_ports
import subprocess
import json
import time
from pathlib import Path
from datetime import datetime, timedelta
from fpdf import FPDF
import re

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def get_base_dir():
    return getattr(sys, 'frozen', False) and sys._MEIPASS or Path(__file__).resolve().parent


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automation Test All Makes")
        self.root.geometry("1000x700")

        # Đường dẫn và cấu hình
        self.base_dir = get_base_dir()
        self.simfile_path = self.base_dir.parent / "Sim 01 41"
        self.database_dir = self.base_dir.parent / "database"
        self.document_path = self.database_dir / "Mode01_41_document.xlsx"
        self.result_path = self.database_dir / "database.json"
        self.result_pdf_path = self.base_dir / "results.pdf"

        self.all_cases = self.base_dir / "Auto_Cases.py"
        self.auto_led_mil = self.base_dir / "Auto_LED_MIL.py"
        self.auto_led_monitor = self.base_dir / "Auto_LED_Monitor.py"

        # Thư mục logs để lưu file log của mỗi case
        self.logs_dir = self.base_dir / "logs"
        self.logs_dir.mkdir(exist_ok=True)

        # Biến giao diện
        self.case_vars = []
        self.com_port = tk.StringVar()
        self.all_cases_var = tk.BooleanVar()
        self.stop_requested = False

        # Dữ liệu từ file Excel
        self.excel_data = self.load_excel_data()
        self.create_widgets()

    def load_excel_data(self):
        try:
            df = pd.read_excel(self.document_path)
            return df
        except FileNotFoundError:
            messagebox.showerror("Error", "File Excel không tìm thấy.")
            return None

    def create_widgets(self):
        # Frame cho danh sách case
        self.case_frame = ttk.LabelFrame(self.root, text="Select Cases", padding=(10, 10))
        self.case_frame.pack(padx=10, pady=10, fill="x")

        # Checkbox cho nhóm chọn lọc
        self.create_group_checkboxes()

        # Canvas cuộn với danh sách case
        self.case_canvas = tk.Canvas(self.case_frame, height=150)
        self.case_canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar = ttk.Scrollbar(self.case_frame, orient="vertical", command=self.case_canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.case_canvas.configure(yscrollcommand=self.scrollbar.set)

        # Frame bên trong canvas để chứa các case checkbox
        self.case_inner_frame = ttk.Frame(self.case_canvas)
        self.case_canvas.create_window((0, 0), window=self.case_inner_frame, anchor="nw")

        # Tạo checkbox cho từng Case từ file Excel
        if self.excel_data is not None:
            case_names = self.excel_data["Case"].dropna().tolist()
            for case_name in case_names:
                case_var = tk.BooleanVar()
                cb = ttk.Checkbutton(self.case_inner_frame, text=case_name, variable=case_var,
                                     command=self.update_selected_cases)
                cb.pack(anchor="w")
                self.case_vars.append((case_name, case_var))

        # Cấu hình cho các điều khiển khác
        self.case_inner_frame.update_idletasks()
        self.case_canvas.config(scrollregion=self.case_canvas.bbox("all"))
        self.case_canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        ttk.Label(self.root, text="Select COM Port:").pack(pady=5)
        self.combobox = ttk.Combobox(self.root, textvariable=self.com_port, postcommand=self.refresh_com_ports)
        self.combobox.pack(pady=5)

        # Nút Scan và hiển thị kết quả
        self.scan_button = ttk.Button(self.root, text="Scan", command=self.run_scan)
        self.scan_button.pack(pady=5)
        self.result_tree = ttk.Treeview(self.root, columns=("Case", "Status", "Relink 4.2s", "Relink 60s", "Result"),
                                        show="headings")
        for col in ["Case", "Status", "Relink 4.2s", "Relink 60s", "Result"]:
            self.result_tree.heading(col, text=col)
        self.result_tree.pack(fill="both", expand=True)

        # Nút mở file PDF kết quả
        self.result_button = ttk.Button(self.root, text="Result", command=self.open_result_pdf)
        self.result_button.pack(pady=10)

    def create_group_checkboxes(self):
        group_frame = ttk.Frame(self.case_frame)
        group_frame.pack(pady=5)

        self.all_cases_checkbox = tk.BooleanVar()
        self.carb_checkbox = tk.BooleanVar()
        self.massachusetts_checkbox = tk.BooleanVar()
        self.no_program_checkbox = tk.BooleanVar()
        self.monitor_icon_checkbox = tk.BooleanVar()

        ttk.Checkbutton(group_frame, text="All Cases", variable=self.all_cases_checkbox,
                        command=self.select_cases_by_name_pattern).pack(side="left")
        ttk.Checkbutton(group_frame, text="CARB", variable=self.carb_checkbox,
                        command=lambda: self.select_cases_by_group("CARB")).pack(side="left")
        ttk.Checkbutton(group_frame, text="Massachusetts", variable=self.massachusetts_checkbox,
                        command=lambda: self.select_cases_by_group("Massachusetts")).pack(side="left")
        ttk.Checkbutton(group_frame, text="No Program", variable=self.no_program_checkbox,
                        command=lambda: self.select_cases_by_group("No Program")).pack(side="left")
        ttk.Checkbutton(group_frame, text="Monitor Icon", variable=self.monitor_icon_checkbox,
                        command=lambda: self.select_cases_by_group("Monitor Icon")).pack(side="left")
        ttk.Button(group_frame, text="Deselect All", command=self.deselect_all_cases).pack(side="left")

    def select_cases_by_name_pattern(self):
        for case_name, case_var in self.case_vars:
            if "Case" in case_name and case_name.endswith(".sim"):
                case_var.set(self.all_cases_checkbox.get())

    def select_cases_by_group(self, group_name):
        for case_name, case_var in self.case_vars:
            if self.excel_data.loc[self.excel_data["Case"] == case_name, "Make"].values[0] == group_name:
                case_var.set(True)

    def deselect_all_cases(self):
        for _, case_var in self.case_vars:
            case_var.set(False)

    def update_selected_cases(self):
        pass

    def on_mousewheel(self, event):
        self.case_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def refresh_com_ports(self):
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.combobox['values'] = port_list

    def get_sim_file_for_case(self, case_name):
        if self.excel_data is not None:
            row = self.excel_data[self.excel_data["Case"] == case_name]
            if not row.empty:
                make_folder = row.iloc[0]["Make"]
                sim_file_name = row.iloc[0]["Case"]
                sim_file_path = self.simfile_path / make_folder / sim_file_name
                if sim_file_path.exists():
                    return sim_file_path
                else:
                    logging.warning(f"SIM file not found: {sim_file_path}")
        return None

    def update_json_file(self, case_name):
        row = self.excel_data[self.excel_data["Case"] == case_name].iloc[0]
        data = {
            "Make": row["Make"],
            "Case": row["Case"],
            "Years": row["Years"],
            "Location": row["Location"],
            "OBD2": "",
            "LED": "",
            "DTC": "",
            "Freeze Frame": "",
            "MIL": "",
            "Monitor Color": ""
        }
        with open(self.result_path, 'w') as file:
            json.dump(data, file, indent=4)

    def run_python_script(self, case_name):
        row = self.excel_data[self.excel_data["Case"] == case_name].iloc[0]
        make = row["Make"]
        if "Case" in case_name:
            script_path = self.all_cases
        elif make in ["CARB", "Massachusetts", "No Program"]:
            script_path = self.auto_led_mil
        elif make == "Monitor Icon":
            script_path = self.auto_led_monitor
        subprocess.run(["python", script_path])

    def run_scan(self):
        selected_cases = [case_name for case_name, var in self.case_vars if var.get()]

        if not selected_cases:
            messagebox.showwarning("Warning", "No case selected.")
            return

        if not self.com_port.get():
            messagebox.showwarning("Warning", "No COM port selected.")
            return

        for row in self.result_tree.get_children():
            self.result_tree.delete(row)

        def process_cases():
            for case_name in selected_cases:
                self.result_tree.insert("", "end", iid=case_name,
                                        values=(case_name, "⏳ Pending", "⏳ Pending", "⏳ Pending"))
                self.close_app("com.innova.passthru")
                time.sleep(3)
                sim_file_path = self.get_sim_file_for_case(case_name)
                if sim_file_path:
                    self.open_app("com.innova.passthru")
                    time.sleep(10)
                    self.update_json_file(case_name)
                    relink_thread = threading.Thread(target=self.monitor_relink)
                    relink_thread.start()
                    self.run_bat_and_monitor(case_name, sim_file_path)
                    self.run_python_script(case_name)
                    self.wait_between_cases()
                else:
                    self.update_result(case_name, "Sim File Not Found", "N/A", "N/A")

        threading.Thread(target=process_cases, daemon=True).start()

    def sanitize_filename(self, case_name):
        return re.sub(r'[\\/*?:"<>| ]', '_', case_name)
    def open_app(self, package_name):
        try:
            subprocess.run(
                ["adb", "shell", "monkey", "-p", package_name, "-c", "android.intent.category.LAUNCHER", "1"],
                check=True
            )
            logging.info(f"{package_name} has been opened successfully.")
            return True
        except subprocess.CalledProcessError as e:
            logging.error(f"Error opening {package_name}: {e}")
            return False

    # Hàm để tắt app
    def close_app(self, package_name):
        try:
            subprocess.run(["adb", "shell", "am", "force-stop", package_name], check=True)
            logging.info(f"{package_name} has been closed successfully.")
            return True
        except subprocess.CalledProcessError as e:
            logging.error(f"Error closing {package_name}: {e}")
            return False
    def stop_running_processes(self):
        subprocess.call(["taskkill", "/F", "/IM", "SimulatorTest.exe"], stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL)
        subprocess.call(["taskkill", "/F", "/IM", "cmd.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Stopped running processes: SimulatorTest.exe, cmd.exe")
        self.stop_requested = True

    def update_bat_file(self, com_port, sim_file):
        bat_file_path = "SimulationWithShowData.bat"
        try:
            bat_content = f'SimulatorTest.exe "{com_port}" "{sim_file}" "showdata"\n'
            with open(bat_file_path, 'w') as bat_file:
                bat_file.write(bat_content)
            logging.info(f"Updated .bat file with COM port {com_port} and SIM file {sim_file}")
        except Exception as e:
            logging.error(f"Error updating .bat file: {e}")

    def run_bat_and_monitor(self, case_name, sim_file):
        sanitized_case_name = self.sanitize_filename(case_name)
        output_file = self.logs_dir / f"log_{sanitized_case_name}.txt"  # Lưu log trong thư mục "logs"
        com_port = self.com_port.get()
        self.update_bat_file(com_port, sim_file)
        self.stop_running_processes()

        command = f'start cmd.exe /c "SimulationWithShowData.bat > {output_file}"'
        subprocess.Popen(command, shell=True)

    def wait_between_cases(self):
        time.sleep(5)

    def monitor_relink(self):
        with open(self.result_path, 'r') as file:
            data = json.load(file)

        if data.get("OBD2", "").lower() == "true":
            print("[INFO] OBD2 is True - starting Relink check")
            start_time = datetime.now()
            end_time = start_time + timedelta(seconds=180)

            match_4_2_sequence = []
            match_60s_sequence = []

            while datetime.now() <= end_time:
                with open(self.result_path, 'r') as file:
                    data = json.load(file)

                if data.get("Relink 4.2s") == "Pass":
                    match_4_2_sequence.append(datetime.now())
                if data.get("Relink 60s") == "Pass":
                    match_60s_sequence.append(datetime.now())

                time.sleep(1)

            self.evaluate_relink_results(match_4_2_sequence, match_60s_sequence)

    def evaluate_relink_results(self, match_4_2_sequence, match_60s_sequence):
        valid_4_2s = len(match_4_2_sequence) >= 10
        valid_60s = len(match_60s_sequence) >= 2

        result_4_2s = "Pass" if valid_4_2s else "Fail"
        result_60s = "Pass" if valid_60s else "Fail"
        return result_4_2s, result_60s

    def update_result(self, case_name, status, result_4_2s="Pending", result_60s="Pending"):
        with open(self.result_path, 'r') as file:
            json_data = json.load(file)

        result = "Pass" if all(
            json_data.get(col, "") == self.excel_data.loc[self.excel_data["Case"] == case_name, col].values[0]
            for col in json_data if col != "OBD2"
        ) else "Fail"

        self.result_tree.item(case_name, values=(case_name, status, result_4_2s, result_60s, result))

    def open_result_pdf(self):
        self.create_pdf_result()
        subprocess.Popen([self.result_pdf_path], shell=True)

    def create_pdf_result(self):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Automation Test Results", ln=True, align="C")
        for item in self.result_tree.get_children():
            case, status, relink_4_2s, relink_60s, result = self.result_tree.item(item, "values")
            pdf.cell(200, 10, txt=f"{case}: Status={status}, Result={result}", ln=True)
        pdf.output(self.result_pdf_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
