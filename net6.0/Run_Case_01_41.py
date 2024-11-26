import sys
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import subprocess
import logging
from pathlib import Path
import time
import serial.tools.list_ports  # Thư viện để lấy danh sách COM

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_base_dir():
    return getattr(sys, 'frozen', False) and sys._MEIPASS or Path(__file__).resolve().parent

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automation Test All Makes")
        self.root.geometry("600x400")

        self.base_dir = get_base_dir()
        self.simfile_path = self.base_dir.parent / "Sim 01 41"
        self.database_dir = self.base_dir.parent / "database"
        self.document_path = self.database_dir / "Mode01_41_document.xlsx"

        self.selected_cases = [tk.BooleanVar() for _ in range(30)]
        self.all_cases_var = tk.BooleanVar()
        self.com_port = tk.StringVar()  # Biến để lưu cổng COM đã chọn

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
        self.case_frame = ttk.LabelFrame(self.root, text="Select Cases", padding=(10, 10))
        self.case_frame.pack(padx=10, pady=10, fill="x")

        for i in range(30):
            ttk.Checkbutton(self.case_frame, text=f"Case {i + 1}", variable=self.selected_cases[i]).grid(
                row=i // 5, column=i % 5, sticky="w"
            )

        self.all_button = ttk.Checkbutton(self.root, text="All", variable=self.all_cases_var,
                                          command=self.select_all_cases)
        self.all_button.pack(pady=5)

        # ComboBox để chọn cổng COM
        ttk.Label(self.root, text="Select COM Port:").pack(pady=5)
        self.combobox = ttk.Combobox(self.root, textvariable=self.com_port, postcommand=self.refresh_com_ports)
        self.combobox.pack(pady=5)

        self.scan_button = ttk.Button(self.root, text="Scan", command=self.run_scan)
        self.scan_button.pack(pady=5)

        self.result_tree = ttk.Treeview(self.root, columns=("Case", "Status"), show="headings")
        self.result_tree.heading("Case", text="Case")
        self.result_tree.heading("Status", text="Status")
        self.result_tree.pack(fill="both", expand=True)

    def select_all_cases(self):
        for var in self.selected_cases:
            var.set(self.all_cases_var.get())

    def refresh_com_ports(self):
        """Cập nhật danh sách cổng COM trong ComboBox."""
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.combobox['values'] = port_list

    def run_scan(self):
        selected_case_indices = [i + 1 for i, var in enumerate(self.selected_cases) if var.get()]

        if not selected_case_indices:
            messagebox.showwarning("Warning", "No case selected.")
            return

        if not self.com_port.get():
            messagebox.showwarning("Warning", "No COM port selected.")
            return

        for row in self.result_tree.get_children():
            self.result_tree.delete(row)

        for case_id in selected_case_indices:
            sim_file_path = self.get_sim_file_for_case(case_id)
            if sim_file_path:
                status = self.run_bat_file_with_sim(case_id, sim_file_path)
                self.result_tree.insert("", "end", values=(f"Case {case_id}", status))
            else:
                self.result_tree.insert("", "end", values=(f"Case {case_id}", "Sim File Not Found"))

    def get_sim_file_for_case(self, case_id):
        if self.excel_data is not None:
            row = self.excel_data[self.excel_data["Case"] == f"Case {case_id}.sim"]
            if not row.empty:
                make_folder = row.iloc[0]["Make"]
                sim_file_name = row.iloc[0]["Case"]
                sim_file_path = self.simfile_path / make_folder / sim_file_name
                if sim_file_path.exists():
                    return sim_file_path
                else:
                    logging.warning(f"SIM file not found: {sim_file_path}")
        return None

    def stop_running_processes(self):
        subprocess.call(["taskkill", "/F", "/IM", "SimulatorTest.exe"], stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL)
        subprocess.call(["taskkill", "/F", "/IM", "cmd.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Stopped running processes: SimulatorTest.exe, cmd.exe")
        self.stop_requested = True

    def update_bat_file(self, com_port, sim_file):
        """Cập nhật nội dung file .bat với COM port và đường dẫn SIM file."""
        bat_file_path = "SimulationWithShowData.bat"
        try:
            # Tạo nội dung mới cho file .bat
            bat_content = f'SimulatorTest.exe "{com_port}" "{sim_file}" "showdata"\n'
            with open(bat_file_path, 'w') as bat_file:
                bat_file.write(bat_content)
            logging.info(f"Updated .bat file with COM port {com_port} and SIM file {sim_file}")
        except Exception as e:
            logging.error(f"Error updating .bat file: {e}")

    def run_bat_file_with_sim(self, case_id, sim_file):
        output_file = f"log_case_{case_id}.txt"
        com_port = self.com_port.get()  # Lấy cổng COM được chọn

        # Cập nhật file .bat trước khi chạy
        self.update_bat_file(com_port, sim_file)

        self.stop_running_processes()
        command = f'start cmd.exe /c "SimulationWithShowData.bat > {output_file}"'

        try:
            process = subprocess.Popen(command, shell=True)
            process.communicate(timeout=60)
            logging.info(f"Case {case_id}: Batch file executed with SIM file {sim_file} and COM port {com_port}")
            # time.sleep(15)
            self.run_all_functions()
            return "Completed"
        except subprocess.TimeoutExpired:
            process.kill()
            logging.warning(f"Case {case_id}: Execution timed out.")
            return "Timeout"
        except Exception as e:
            logging.error(f"Case {case_id}: Error executing batch file: {e}")
            return "Error"

    def run_all_functions(self):
        while True:
            time.sleep(1)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
