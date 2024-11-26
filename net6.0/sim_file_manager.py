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


class DropdownWithScrollbar(tk.Frame):
    def __init__(self, parent, group_name, case_vars, max_height=300, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.group_name = group_name
        self.case_vars = case_vars

        # Button to toggle drop-down
        self.toggle_button = ttk.Button(self, text=f"{self.group_name} ▼", command=self.toggle_dropdown)
        self.toggle_button.pack(fill="x")

        # Frame for the checkboxes (with a scrollbar)
        self.check_frame = ttk.Frame(self)
        self.canvas = tk.Canvas(self.check_frame, height=max_height)
        self.scrollbar = ttk.Scrollbar(self.check_frame, orient="vertical", command=self.canvas.yview)
        self.inner_frame = ttk.Frame(self.canvas)

        # Create window for inner frame on canvas
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.canvas.config(yscrollcommand=self.scrollbar.set)

        # Populate checkboxes
        for case_name, var in self.case_vars:
            cb = ttk.Checkbutton(self.inner_frame, text=case_name, variable=var)
            cb.pack(anchor="w")

        # Configure scrolling region and initial visibility
        self.inner_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.check_frame.pack_forget()
        self.check_frame_visible = False

    def toggle_dropdown(self):
        if self.check_frame_visible:
            self.check_frame.pack_forget()
            self.toggle_button.config(text=f"{self.group_name} ▼")
        else:
            self.check_frame.pack(fill="x")
            self.toggle_button.config(text=f"{self.group_name} ▲")
        self.check_frame_visible = not self.check_frame_visible


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automation Test All Makes")
        self.root.geometry("1000x800")

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
        # Main container frame with two columns
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Left column frame for case selection
        self.case_frame = ttk.LabelFrame(main_frame, text="Select Cases", padding=(10, 10))
        self.case_frame.grid(row=0, column=0, sticky="nswe", padx=(0, 10))

        # Checkbox group selection
        self.create_group_checkboxes()

        # Create dropdowns for each group
        if self.excel_data is not None:
            case_names = self.excel_data["Case"].dropna().tolist()
            make_names = self.excel_data["Make"].dropna().tolist()

            # Group cases by 'Make' column and create the "All Cases" group based on the rule
            grouped_cases = {
                "All Cases": [],
                "CARB": [],
                "Massachusetts": [],
                "No Program": [],
                "Monitor Icon": []
            }

            for case_name, make_name in zip(case_names, make_names):
                var = tk.BooleanVar()
                # Rule for "All Cases": any case name containing the word "Case"
                if "Case" in case_name:
                    grouped_cases["All Cases"].append((case_name, var))

                # Add to specific groups based on the "Make" column
                if make_name in grouped_cases:
                    grouped_cases[make_name].append((case_name, var))

                # Add to case_vars for global access
                self.case_vars.append((case_name, var))

            # Create dropdowns for each pre-defined group
            for group_name, cases in grouped_cases.items():
                if cases:  # Only create dropdown if there are cases
                    dropdown = DropdownWithScrollbar(self.case_frame, group_name, cases)
                    dropdown.pack(fill="x", pady=5)

        # Right column frame for Control options (Select COM Port, Scan, Result)
        control_frame = ttk.LabelFrame(main_frame, text="Control", padding=(10, 10))
        control_frame.grid(row=0, column=1, sticky="nswe")

        # COM Port selection within the Control frame
        ttk.Label(control_frame, text="Select COM Port:").pack(pady=(0, 5))
        self.combobox = ttk.Combobox(control_frame, textvariable=self.com_port, postcommand=self.refresh_com_ports)
        self.combobox.pack(fill="x", padx=10, pady=(0, 10))

        # Scan and Result buttons within the Control frame
        self.scan_button = ttk.Button(control_frame, text="Scan", command=self.run_scan)
        self.scan_button.pack(fill="x", padx=10, pady=5)

        self.result_button = ttk.Button(control_frame, text="Result", command=self.open_result_pdf)
        self.result_button.pack(fill="x", padx=10, pady=5)

        # Frame for the results table with scrollbar
        result_frame = ttk.Frame(self.root)
        result_frame.pack(fill="both", expand=True, pady=(10, 0))

        # Result Treeview with Scrollbar
        self.result_tree = ttk.Treeview(result_frame, columns=("Case", "Status", "Relink 4.2s", "Relink 60s", "Result"),
                                        show="headings")
        for col in ["Case", "Status", "Relink 4.2s", "Relink 60s", "Result"]:
            self.result_tree.heading(col, text=col)
        self.result_tree.pack(side="left", fill="both", expand=True)

        # Vertical scrollbar for result_tree
        result_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        result_scrollbar.pack(side="right", fill="y")
        self.result_tree.configure(yscrollcommand=result_scrollbar.set)

        # Adjust layout weights to ensure left and right columns expand proportionally
        main_frame.columnconfigure(0, weight=3)  # Left side (case selection) gets more space
        main_frame.columnconfigure(1, weight=1)  # Right side (Control frame) takes less space

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
            "LED": "Null",
            "DTC": "Null",
            "Freeze Frame": "Null",
            "MIL": "Null",
            "Monitor Color": "Null"
        }
        with open(self.result_path, 'w') as file:
            json.dump(data, file, indent=4)

    def run_python_script(self, case_name):
        logging.debug(f"Starting run_python_script for case: {case_name}")
        row = self.excel_data[self.excel_data["Case"] == case_name].iloc[0]
        make = row["Make"]

        if "Case" in case_name:
            logging.debug("Detected 'Case' in case_name. Initializing relink thread.")

            def monitor_relink_with_timeout():
                logging.debug("Waiting 20 seconds before running monitor_relink.")
                time.sleep(100)
                logging.debug("Running monitor_relink.")
                self.monitor_relink(case_name)
                logging.debug("monitor_relink running for 180 seconds.")
                # time.sleep(180)

            relink_thread = threading.Thread(target=monitor_relink_with_timeout)
            relink_thread.start()

            logging.debug("Running all_cases script.")
            script_path = self.all_cases
            subprocess.run(["python", script_path])

            logging.debug("Waiting for relink_thread to complete.")
            relink_thread.join()
            logging.debug("relink_thread completed.")

        elif make in ["CARB", "Massachusetts", "No Program"]:
            logging.debug("Running auto_led_mil script.")
            script_path = self.auto_led_mil
            subprocess.run(["python", script_path])
        elif make == "Monitor Icon":
            logging.debug("Running auto_led_monitor script.")
            script_path = self.auto_led_monitor
            subprocess.run(["python", script_path])

        logging.debug("Completed run_python_script.")
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

    def run_scan(self):
        selected_cases = [case_name for case_name, var in self.case_vars if var.get()]

        if not selected_cases:
            messagebox.showwarning("Warning", "No case selected.")
            return

        if not self.com_port.get():
            messagebox.showwarning("Warning", "No COM port selected.")
            return

        # Xóa kết quả cũ khỏi bảng kết quả
        for row in self.result_tree.get_children():
            self.result_tree.delete(row)

        def process_cases():
            for case_name in selected_cases:
                # Cập nhật trạng thái case thành "Pending" sử dụng lambda để tránh lỗi từ after()
                self.root.after(0, lambda name=case_name: self.result_tree.insert("", "end", iid=name,
                                                                                  values=(
                                                                                      name, "⏳ Pending", "⏳ Pending",
                                                                                      "⏳ Pending", "⏳ Pending")))

                self.close_app("com.innova.passthru")
                time.sleep(3)
                sim_file_path = self.get_sim_file_for_case(case_name)
                if sim_file_path:
                    self.open_app("com.innova.passthru")
                    time.sleep(10)
                    self.update_json_file(case_name)

                    self.run_bat_and_monitor(case_name, sim_file_path)

                    if "Case" in case_name:
                        self.run_python_script(case_name)
                    else:
                        self.run_python_script(case_name)
                        self.root.after(0, lambda name=case_name: self.update_result(name, "✅ Completed", "N/A", "N/A"))

                    self.wait_between_cases()
                else:
                    self.root.after(0,
                                    lambda name=case_name: self.update_result(name, "Sim Not Found", "N/A", "N/A"))

        # Chạy process_cases trong một thread để tránh khóa UI
        threading.Thread(target=process_cases, daemon=True).start()

    def monitor_relink(self, case_name):
        sanitized_case_name = self.sanitize_filename(case_name)
        output_file = self.logs_dir / f"log_{sanitized_case_name}.txt"
        logging.info(f"Starting monitor_relink for case: {case_name}, log file: {output_file}")

        try:
            start_time = datetime.now()
            end_time = start_time + timedelta(seconds=180)
            last_read_position = 0

            patterns_4_2s = ["08 02 01 0C", "08 02 01 41"]
            patterns_60s = ["08 02 01", "08 03 02", "08 01 03", "08 01 07", "08 01 0A"]

            match_4_2_sequence = []
            match_60s_sequence = []
            current_4_2_index = 0
            current_60_index = 0

            while datetime.now() <= end_time:
                if not output_file.exists():
                    logging.warning(f"Log file does not exist: {output_file}")
                    time.sleep(1)
                    continue

                with open(output_file, "r") as file:
                    file.seek(last_read_position)
                    new_lines = file.readlines()
                    last_read_position = file.tell()

                for line in new_lines:
                    logging.debug(f"Read line: {line.strip()}")
                    line = line.strip()
                    if line:
                        if patterns_4_2s[current_4_2_index] in line:
                            match_4_2_sequence.append(datetime.now())
                            logging.info(f"4.2s pattern matched: {patterns_4_2s}")
                            current_4_2_index += 1
                            if current_4_2_index == len(patterns_4_2s):
                                current_4_2_index = 0

                        if patterns_60s[current_60_index] in line:
                            match_60s_sequence.append(datetime.now())
                            logging.info(f"4.2s pattern matched: {patterns_60s}")
                            current_60_index += 1
                            if current_60_index == len(patterns_60s):
                                current_60_index = 0
                #
                # logging.info(f"Progress: {len(match_4_2_sequence)} matches for 4.2s patterns, "
                #              f"{len(match_60s_sequence)} matches for 60s patterns.")
                print(f"Monitor Relink Progress - {case_name}: "
                      f"{len(match_4_2_sequence)} matches for 4.2s patterns, "
                      f"{len(match_60s_sequence)} matches for 60s patterns.")

                # Sleep for a short interval to avoid busy-waiting
                time.sleep(1)

            logging.info(f"Completed monitor_relink for case: {case_name}")
            print(f"Monitor Relink Completed - {case_name}: "
                  f"{len(match_4_2_sequence)} matches for 4.2s patterns, "
                  f"{len(match_60s_sequence)} matches for 60s patterns.")
            valid_4_2s = len(match_4_2_sequence) >= len(patterns_4_2s) * 10
            valid_60s = len(match_60s_sequence) >= len(patterns_60s) * 2

            result_4_2s = "✅ Passed" if valid_4_2s else "❌ Failed"
            result_60s = "✅ Passed" if valid_60s else "❌ Failed"

            # Cập nhật UI từ luồng chính sau khi hoàn tất
            self.root.after(0, self.update_result, case_name, "✅ Completed", result_4_2s, result_60s)
        except Exception as e:
            logging.error(f"Error in monitor_relink: {e}")

    def update_result(self, case_name, status, result_4_2s, result_60s):
        with open(self.result_path, 'r') as file:
            json_data = json.load(file)

        # Kiểm tra kết quả so sánh JSON với dữ liệu Excel
        result = "✅ Passed" if all(
            json_data.get(col, "") == self.excel_data.loc[self.excel_data["Case"] == case_name, col].values[0]
            for col in json_data if col != "OBD2"
        ) else "❌ Failed"

        # Sử dụng after để cập nhật giao diện từ luồng chính
        self.root.after(0, self._update_result_on_main_thread, case_name, status, result_4_2s, result_60s, result)

    def _update_result_on_main_thread(self, case_name, status, result_4_2s, result_60s, result):
        self.result_tree.item(case_name, values=(case_name, status, result_4_2s, result_60s, result))

    def open_result_pdf(self):
        self.create_pdf_result()
        subprocess.Popen([self.result_pdf_path], shell=True)

    def create_pdf_result(self):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)  # Tự động chuyển trang với lề 15mm
        pdf.set_font("Arial", size=12)

        # Title
        pdf.cell(0, 10, txt="Automation Test Results", ln=True, align="C")

        # Define left and right margins for centered alignment
        page_width = 210  # A4 page width in mm
        table_width = 200  # Total width of the table (sum of column widths)
        left_margin = (page_width - table_width) / 2
        pdf.set_left_margin(left_margin)
        pdf.set_right_margin(left_margin)
        pdf.set_x(left_margin)

        # Column headers and widths
        headers = ["Case", "Status", "Relink 4.2s", "Relink 60s", "Result"]
        column_widths = [100, 25, 25, 25, 25]  # Widths for each column

        # Draw header row
        pdf.set_font("Arial", style="B", size=10)
        for i, header in enumerate(headers):
            pdf.cell(column_widths[i], 10, header, border=1, align="C")
        pdf.ln()

        # Set font for data rows
        pdf.set_font("Arial", size=10)

        for item in self.result_tree.get_children():
            case, status, relink_4_2s, relink_60s, result = self.result_tree.item(item, "values")

            # Replace icons with descriptive text
            status_text = status.replace("✅ Passed", "Passed").replace("❌ Failed", "Failed").replace("⚠️",
                                                                                                     "Error").replace(
                "⏳ Pending", "Pending").replace("✅ Completed", "Completed")
            relink_4_2s_text = relink_4_2s.replace("✅ Passed", "Passed").replace("❌ Failed", "Failed").replace(
                "⏳ Pending", "Pending")
            relink_60s_text = relink_60s.replace("✅ Passed", "Passed").replace("❌ Failed", "Failed").replace(
                "⏳ Pending", "Pending")
            result_text = result.replace("✅ Passed", "Passed").replace("❌ Failed", "Failed").replace("⏳ Pending",
                                                                                                     "Pending")

            # Remove any unnecessary newlines or formatting issues in the "Case" text
            case = case.replace('\n', ' ')  # Loại bỏ xuống dòng không mong muốn trong "Case"

            # Calculate dynamic row height
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            pdf.multi_cell(column_widths[0], 10, case, border=1)

            # Calculate the current Y position after multi_cell
            y_end = pdf.get_y()
            row_height = y_end - y_start  # Height of the row

            # Return to the starting position for the next columns
            pdf.set_xy(x_start + column_widths[0], y_start)

            # Fill other columns with the same row height
            pdf.cell(column_widths[1], row_height, status_text, border=1)
            pdf.cell(column_widths[2], row_height, relink_4_2s_text, border=1)
            pdf.cell(column_widths[3], row_height, relink_60s_text, border=1)
            pdf.cell(column_widths[4], row_height, result_text, border=1)

            # Move to the next line at the correct position
            pdf.set_y(y_end)

            # Check if a new page is needed
            if pdf.get_y() + row_height > pdf.page_break_trigger:
                pdf.add_page()

        # Output the PDF file
        pdf.output(self.result_pdf_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
