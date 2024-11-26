import csv
import json
import logging
import os
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import pandas as pd
import psutil
import serial.tools.list_ports
from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.common.appiumby import AppiumBy
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException, \
    StaleElementReferenceException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def get_base_dir():
    return getattr(sys, 'frozen', False) and sys._MEIPASS or Path(__file__).resolve().parent


class Config:
    def __init__(self):
        self.base_dir = get_base_dir()
        self.database_dir = self.base_dir.parent / "database"
        self.setting_path = self.database_dir / "setting.json"
        self.simfile_path = self.base_dir.parent / "Sim files"
        self.bat_file_1 = self.base_dir / "SimulationWithShowData.bat"
        self.bat_file_2 = self.base_dir / "SimulationWithShowData2.bat"
        self.all_functions = self.base_dir.parent / "All_Functions.py"
        self.nws_live_data_functions = self.base_dir.parent / "NWS_LiveData.py"
        self.obd2_10modes = self.base_dir.parent / "OBD2_10Modes.py"
        self.obd2_livedata = self.base_dir.parent / "OBD2_LiveData.py"
        self.nws_dtcs = self.base_dir.parent / "NWS_DTCs.py"
        self.txt_path = self.base_dir.parent / "VIN Decode.txt"
        self.net_6_dir = self.base_dir.parent / "net6.0"
        self.auto_source = self.net_6_dir / "AutoTest_AllMakes.py"
        self.ensure_directories_exist()

    def ensure_directories_exist(self):
        directories = [self.database_dir, self.simfile_path]
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            logging.info(f"Directory exists or created: {directory}")

config = Config()


class DeviceManager:
    def __init__(self, desired_caps, appium_server_url):
        self.desired_caps = desired_caps
        self.appium_server_url = appium_server_url
        self.driver = None

    def initialize_driver(self):
        try:
            options = UiAutomator2Options()
            options.load_capabilities(self.desired_caps)
            driver = webdriver.Remote(self.appium_server_url, options=options)
            logging.info("Appium driver initialized.")
            return driver
        except Exception as e:
            logging.error(f"Failed to initialize Appium driver: {e}")
            return None

    def restart_device(self):
        try:
            subprocess.run(["adb", "reboot"], check=True)
            logging.info("Device is restarting...")
            if not self.wait_for_device_to_be_ready():
                logging.warning("Device did not become ready in time.")
                return False
            logging.info("Device is fully booted and ready.")
            return True
        except subprocess.CalledProcessError:
            logging.exception("Failed to restart device")
            return False

    def wait_for_device_to_be_ready(self):
        for _ in range(60):
            result = subprocess.run(['adb', 'shell', 'getprop', 'sys.boot_completed'], capture_output=True, text=True)
            if result.stdout.strip() == '1' and self.check_device_connection():
                time.sleep(10)
                logging.info("Device is fully booted and ready.")
                return True
            logging.info("Waiting for device to be ready...")
            time.sleep(10)
        return False

    def handle_app_crash(self):
        logging.error("App crash detected. Restarting driver.")
        self.restart_uiautomator2_server()

    def restart_uiautomator2_server(self):
        if self.driver is not None:
            self.driver.quit()
        self.driver = self.initialize_driver()

    def check_device_connection(self):
        result = subprocess.run(['adb', 'devices'], capture_output=True, text=True)
        return "device" in result.stdout and "unauthorized" not in result.stdout

    def restart_app(self, package_name, check_ui_ready=False):
        try:
            subprocess.run(["adb", "shell", "am", "force-stop", package_name], check=True)
            subprocess.run(
                ["adb", "shell", "monkey", "-p", package_name, "-c", "android.intent.category.LAUNCHER", "1"],
                check=True)
            logging.info(f"{package_name} has been restarted successfully.")

            if check_ui_ready and not self.wait_for_app_to_be_ready(package_name):
                logging.warning(f"{package_name} did not fully load in time.")
                return False

            return True

        except Exception as e:
            logging.error(f"Error restarting {package_name}: {e}")
            return False

    def wait_for_app_to_be_ready(self, package_name, timeout=120):
        try:
            self.driver = self.initialize_driver()

            if not self.driver:
                logging.error("Appium driver is not properly initialized.")
                return False

            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((AppiumBy.XPATH, "/hierarchy/android.widget.FrameLayout"))
            )
            logging.info(f"{package_name} is fully loaded and ready.")
            return True
        except TimeoutException:
            logging.warning(f"App {package_name} did not load the expected element in time.")
            return False


class SimFileManager:
    def __init__(self, sim_files_path):
        self.sim_files_path = sim_files_path
        self.processes = []
        self.output_files = []
        self.stop_requested = False
    def update_bat_file(self, bat_file, com_port, folder_path, sim_file):
        try:
            com_port = com_port.split()[0]
            with open(bat_file, 'r') as file:
                content = file.read()
            parts = content.split('"')
            parts[1] = com_port
            parts[3] = str(Path(folder_path) / sim_file)
            new_content = '"'.join(parts)
            with open(bat_file, 'w') as file:
                file.write(new_content)
            logging.info(f"Updated {bat_file} with COM port {com_port} and SIM file {sim_file}")
        except PermissionError:
            logging.exception(f"Unable to write to file: {bat_file}")
            messagebox.showerror("Permission Error", f"Unable to write to file: {bat_file}")
        except Exception:
            logging.exception(f"An error occurred while updating {bat_file}")
            messagebox.showerror("Error", f"An error occurred while updating {bat_file}")
    def stop_running_processes(self):
        subprocess.call(["taskkill", "/F", "/IM", "SimulatorTest.exe"], stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL)
        subprocess.call(["taskkill", "/F", "/IM", "cmd.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Stopped running processes: SimulatorTest.exe, cmd.exe")
        self.stop_requested = True

    def run_bat_file(self, bat_file, output_file):
        with open(output_file, 'w') as file:
            pass

        command = f'start cmd.exe /c "{bat_file} > {output_file}"'
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=True,
            text=True
        )
        self.processes.append(process)
        self.output_files.append(output_file)

    def run_bat_files(self, bat_file1, bat_file2):
        self.stop_running_processes()
        time.sleep(3)
        output_file1 = "log_batch_1.txt"
        output_file2 = "log_batch_2.txt"

        with open(output_file1, 'w'), open(output_file2, 'w'):
            pass

        self.run_bat_file(bat_file1, output_file1)
        self.run_bat_file(bat_file2, output_file2)
        logging.info(f"Batch files running: {bat_file1}, {bat_file2}")
        return output_file1, output_file2

    def read_output(self, output_file):
        with open(output_file, 'r') as file:
            return file.read()

    def wait_for_completion(self, timeout=600):
        start_time = time.time()
        while time.time() - start_time < timeout:
            if self.stop_requested:
                logging.info("Stop requested, exiting wait_for_completion.")
                break
            all_done = True
            for proc in self.processes:
                if proc.poll() is None:
                    all_done = False
                    break
            if all_done:
                break
            time.sleep(1)

    def check_bat_file_output(self, output_file1, output_file2=None):
        output_files = [output_file1, output_file2] if output_file2 else [output_file1]
        for output_file in output_files:
            for _ in range(5):
                output = self.read_output(output_file)
                if "Press ESC to exit" in output:
                    lines_after_esc = output.split("Press ESC to exit")[-1].strip().splitlines()
                    valid_com_count = 0

                    for line in lines_after_esc:
                        if line.startswith("COM"):
                            if "[0 ms]" not in line:
                                valid_com_count += 1
                            if valid_com_count >= 10:
                                return True
                time.sleep(10)
        return True

    def validate_sim_file(self, sim_file, data):
        for record in data:
            expected_prefix = f"{record['Year']} {record['Make']} {record['Model']} {record['Engine']}"
            if all(part in sim_file for part in expected_prefix.split()):
                logging.info(f"SIM file {sim_file} is valid with expected prefix {expected_prefix}")
                return True
        logging.warning(f"SIM file {sim_file} is not valid")
        return False


class App:
    desired_caps = {
        "platformName": "android",
        "deviceName": "3c000c4d7641c861eda",
        "automationName": "uiautomator2",
        'newCommandTimeout': 0,
        'skipServerInstallation': True,
        'uiautomator2ServerLaunchTimeout': 120000
    }
    appium_server_url = 'http://localhost:4723'

    config = Config()

    def __init__(self, root):
        self.root = root
        self.root.title("Automation Test All Makes")
        self.root.geometry("500x500")

        # Setting up paths and objects
        self.base_dir = get_base_dir()
        sim_files_path = config.simfile_path
        icon_path = config.base_dir / "Logo.ico"
        self.root.iconbitmap(icon_path)
        self.device_manager = DeviceManager(self.desired_caps, self.appium_server_url)
        self.sim_file_manager = SimFileManager(sim_files_path)

        self.create_widgets()
        self.update_excel_file_list()
        self.excel_file_combo.bind("<<ComboboxSelected>>", self.update_selected_excel)
        self.update_folder_combobox()

        self.scanning = False
        self.loading = False
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        style = ttk.Style()
        style.configure("Green.TButton", background="green")
        style.configure("Yellow.TButton", background="yellow")
        self.appium_process = None
    def create_widgets(self):
        # Settings Group Frame
        self.settings_frame = ttk.LabelFrame(self.root, text="Settings", padding=(10, 10))
        self.settings_frame.pack(padx=10, pady=10, fill="x")

        # Excel file selection in Settings group
        self.excel_file_combo = self.create_select_document_combobox(self.settings_frame, "Select Excel File:")
        self.add_file_button = ttk.Button(self.settings_frame, text="Add File", command=self.add_excel_file)
        self.add_file_button.grid(row=0, column=2, padx=5, pady=5)

        # HID1, HID2, Make in Settings group
        self.combobox_hid1 = self.create_combobox(self.settings_frame, "HID 1 (COM Port):")
        self.combobox_hid2 = self.create_combobox(self.settings_frame, "HID 2 (COM Port):")
        self.folder_combobox = self.create_combobox(self.settings_frame, "Make:")

        # Bind the combobox click event to refresh COM ports when clicked
        self.combobox_hid1.bind("<Button-1>", lambda event: self.scan_ports())
        self.combobox_hid2.bind("<Button-1>", lambda event: self.scan_ports())

        # Function Group Frame
        self.function_frame = ttk.LabelFrame(self.root, text="Function", padding=(10, 10))
        self.function_frame.pack(padx=10, pady=10, fill="x")

        # Checkbuttons for functions
        self.create_function_checkbuttons()

        # Scan Folder and Scan All buttons
        self.scan_folder_button = ttk.Button(self.root, text="Scan Folder", command=self.scan_folder)
        self.scan_folder_button.pack(pady=5)

        self.scan_all_button = ttk.Button(self.root, text="Scan All", command=self.scan_all)
        self.scan_all_button.pack(pady=5)
        # Make them larger and colored
        self.scan_folder_button.config(width=30)
        self.scan_folder_button.config(style="Green.TButton")  # Green for Scan Folder

        self.scan_all_button.config(width=30)
        self.scan_all_button.config(style="Yellow.TButton")  # Yellow for Scan All
        # Progress bar
        self.progress_bar_frame = ttk.Frame(self.root)
        self.progress_bar_frame.pack(pady=10)
        self.progress_bar = ttk.Progressbar(self.progress_bar_frame, orient='horizontal', mode='indeterminate',
                                            length=400)
        self.progress_bar.pack()

    def create_combobox(self, parent, label_text):
        """Helper to create a combobox with a label."""
        row = len(parent.grid_slaves()) // 2  # To place new row
        label = ttk.Label(parent, text=label_text)
        label.grid(row=row, column=0, sticky="w", padx=5, pady=5)
        combobox = ttk.Combobox(parent, state="readonly", width=30)
        combobox.grid(row=row, column=1, padx=5, pady=5)
        return combobox

    def create_select_document_combobox(self, parent, label_text):
        """Helper to create a select document combobox."""
        row = len(parent.grid_slaves()) // 2  # To place new row
        label = ttk.Label(parent, text=label_text)
        label.grid(row=row, column=0, sticky="w", padx=5, pady=5)
        combobox = ttk.Combobox(parent, state="readonly", width=40)
        combobox.grid(row=row, column=1, padx=5, pady=5)
        return combobox

    def create_function_checkbuttons(self):
        """Helper to create all function checkbuttons with fixed spacing."""
        # Define checkbutton variables
        self.all_var = tk.BooleanVar()
        self.restart_device_var = tk.BooleanVar()
        self.obd2_10modes_var = tk.BooleanVar()
        self.obd2_livedata_var = tk.BooleanVar()
        self.obd2_led_logic_var = tk.BooleanVar()
        self.nws_dtcs_var = tk.BooleanVar()
        self.nws_livedata_var = tk.BooleanVar()

        # Create checkbuttons
        self.create_checkbutton(self.function_frame, "OBD2 10 Modes", self.obd2_10modes_var, row=0, column=0)
        self.create_checkbutton(self.function_frame, "OBD2 Live Data", self.obd2_livedata_var, row=0, column=1)
        # self.create_checkbutton(self.function_frame, "OBD2 ($01/$41/LED Logic)", self.obd2_led_logic_var, row=0,column=2)

        self.create_checkbutton(self.function_frame, "NWS DTCs", self.nws_dtcs_var, row=1, column=0)
        self.create_checkbutton(self.function_frame, "NWS Live Data", self.nws_livedata_var, row=1, column=1)

        self.create_checkbutton(self.function_frame, "All Function", self.all_var, row=2, column=0)
        self.create_checkbutton(self.function_frame, "Restart Device", self.restart_device_var, row=2, column=1)



    def create_checkbutton(self, parent, text, variable, row, column):
        """Helper to create a checkbutton and place it on a grid."""
        checkbutton = ttk.Checkbutton(parent, text=text, variable=variable)
        checkbutton.grid(row=row, column=column, padx=10, pady=5, sticky="w")

    def scan_ports(self):
        """Scan available COM ports and update HID comboboxes."""
        ports = serial.tools.list_ports.comports()
        port_list = [f"{port.device} - {port.description}" for port in ports]  # Add port description
        self.combobox_hid1['values'] = port_list
        self.combobox_hid2['values'] = port_list
        logging.info(f"Available COM ports: {port_list}")

    def update_folder_combobox(self):
        """Update Make combobox with available folders."""
        try:
            folders = [f for f in self.sim_file_manager.sim_files_path.iterdir() if f.is_dir()]
            self.folder_combobox['values'] = [f.name for f in folders]
            logging.info(f"Available folders: {[f.name for f in folders]}")
        except FileNotFoundError:
            logging.exception("Error accessing sim files directory")
            messagebox.showerror("Error", "Could not access SIM files directory.")

    def find_excel_files(self):
        """Find Excel files in the database directory."""
        excel_directory = self.base_dir.parent / "database"
        return [file for file in os.listdir(excel_directory) if file.endswith(".xlsx")]

    def update_excel_file_list(self):
        """Update Excel combobox with available files."""
        excel_files = self.find_excel_files()
        if excel_files:
            self.excel_file_combo['values'] = excel_files

    def add_excel_file(self):
        """Add new Excel file."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            destination = os.path.join(f"{self.base_dir.parent}/database", os.path.basename(file_path))
            shutil.copy(file_path, destination)
            self.update_excel_file_list()
            self.excel_file_combo.set(os.path.basename(file_path))
            self.update_setting_file(file_path)

    def update_setting_file(self, file_path):
        """Update setting.json with the selected Excel file."""
        try:
            with open(config.setting_path, 'r') as json_file:
                settings = json.load(json_file)
            settings.update({'test_document': os.path.basename(file_path)})
            with open(config.setting_path, 'w') as json_file:
                json.dump(settings, json_file, indent=4)
            logging.info(f"Updated setting.json with test_document: {os.path.basename(file_path)}")
        except (FileNotFoundError, PermissionError):
            logging.exception("Error updating setting.json")
        except Exception:
            logging.exception("Unexpected error updating setting.json")

    def update_selected_excel(self, event=None):
        """Update settings when an Excel file is selected."""
        self.selected_excel_file = self.excel_file_combo.get()
        if self.selected_excel_file:
            self.update_setting_file(self.selected_excel_file)

    def stop_scan(self):
        self.sim_file_manager.stop_requested = True
        self.scanning = False

        self.sim_file_manager.stop_running_processes()

        logging.info("Stopped all current scan processes.")

    def scan_folder(self):
        if self.loading:
            return
        self.stop_scan()

        self.start_loading_animation()
        self.sim_file_manager.stop_requested = False
        self.scanning = True

        selected_folder = self.folder_combobox.get()
        if not selected_folder:
            messagebox.showwarning("Warning", "Please select a folder.")
            logging.warning("No folder selected.")
            return

        folder_path = self.sim_file_manager.sim_files_path / selected_folder
        sim_files = [f for f in folder_path.iterdir() if f.suffix == '.sim' and not f.name.endswith('.correct.sim')]
        logging.info(f"Scanning folder: {selected_folder} with SIM files: {sim_files}")
        data = self.load_data_excel(selected_folder)
        data = self.remove_duplicates(data)
        logging.info(f"Loaded data from sheet {selected_folder}: {data}")
        valid_sim_files = [f for f in sim_files if self.sim_file_manager.validate_sim_file(f.name, data)]
        self.process_sim_files(selected_folder, folder_path, valid_sim_files, data)

    def scan_all(self):
        if self.loading:
            return

        self.stop_scan()

        self.start_loading_animation()
        self.sim_file_manager.stop_requested = False
        self.scanning = True

        self.scan_folder_button.config(state=tk.DISABLED)
        self.scan_all_button.config(state=tk.DISABLED)

        def run_all_folders():
            for folder in self.folder_combobox['values']:
                if not self.scanning:
                    break
                folder_path = self.sim_file_manager.sim_files_path / folder
                sim_files = [f for f in folder_path.iterdir() if
                             f.suffix == '.sim' and not f.name.endswith('.correct.sim')]
                logging.info(f"Scanning folder: {folder} with SIM files: {sim_files}")
                data = self.load_data_excel(folder)
                data = self.remove_duplicates(data)
                logging.info(f"Loaded data from sheet: {data}")
                valid_sim_files = [f for f in sim_files if self.sim_file_manager.validate_sim_file(f.name, data)]
                self.process_sim_files_sequentially(folder, folder_path, valid_sim_files, data)

            self.scan_folder_button.config(state=tk.NORMAL)
            self.scan_all_button.config(state=tk.NORMAL)
            self.stop_loading_animation()
            self.scanning = False

        threading.Thread(target=run_all_folders).start()

    def start_loading_animation(self):
        self.loading = True
        self.progress_bar.start(10)

    def on_closing(self):
        self.stop_scan()
        if self.appium_process:
            try:
                self.appium_process.terminate()
                self.appium_process.wait()
            except Exception:
                logging.exception("Error terminating Appium process")
        current_pid = os.getpid()
        for proc in psutil.process_iter(['pid', 'ppid']):
            if proc.info['ppid'] == current_pid:
                try:
                    proc.terminate()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        self.root.destroy()

    def stop_loading_animation(self):
        self.loading = False
        self.progress_bar.stop()

    def process_sim_files_sequentially(self, selected_folder, folder_path, sim_files, data):
        com1 = self.combobox_hid1.get()
        com2 = self.combobox_hid2.get()
        if not com1 or not com2:
            messagebox.showwarning("Warning", "Please select both COM ports.")
            logging.warning("COM ports not selected.")
            return
        bat_file1 = config.bat_file_1
        bat_file2 = config.bat_file_2

        i = 0
        while i < len(sim_files):
            if not self.scanning:
                break

            sim_file1 = sim_files[i]
            sim_file2 = sim_files[i + 1] if i + 1 < len(sim_files) else None

            if sim_file2 and self.are_sim_files_matching(sim_file1.name, sim_file2.name):
                self.sim_file_manager.update_bat_file(bat_file1, com1, folder_path, sim_file1.name)
                self.sim_file_manager.update_bat_file(bat_file2, com2, folder_path, sim_file2.name)
                output_file1, output_file2 = self.sim_file_manager.run_bat_files(bat_file1, bat_file2)
                i += 2
            else:
                self.sim_file_manager.update_bat_file(bat_file1, com1, folder_path, sim_file1.name)
                output_file1 = "log_batch_1.txt"
                self.sim_file_manager.run_bat_file(bat_file1, output_file1)
                output_file2 = None
                i += 1

            vin_updated = False
            for record in data:
                expected_prefix = f"{record['Year']} {record['Make']} {record['Model']} {record['Engine']}"
                if all(part in sim_file1.name for part in expected_prefix.split()):
                    vin = record['VIN']
                    self.update_setting(vin, selected_folder)
                    vin_updated = True
                    break

            if not vin_updated:
                logging.error(f"No matching record found for SIM file: {sim_file1.name}")

            time.sleep(3)

            if self.restart_device_var.get():
                self.device_manager.restart_device()

            while not self.device_manager.check_device_connection():
                if not self.scanning:
                    break
                logging.info("Device not connected. Waiting for 10 seconds...")
                time.sleep(10)

            self.device_manager.restart_app('com.innova.passthru')

            if not self.sim_file_manager.check_bat_file_output(output_file1, output_file2):
                logging.error("Failed to connect with bat files after app restart.")
                continue
            time.sleep(3)
            self.run_each_VIN(selected_folder)
            self.sim_file_manager.stop_running_processes()

        logging.info("Completed processing all SIM files.")

    def process_sim_files(self, selected_folder, folder_path, sim_files, data):
        com1 = self.combobox_hid1.get()
        com2 = self.combobox_hid2.get()
        if not com1 or not com2:
            messagebox.showwarning("Warning", "Please select both COM ports.")
            logging.warning("COM ports not selected.")
            return
        bat_file1 = config.bat_file_1
        bat_file2 = config.bat_file_2

        def run_all():
            logging.info(f"Processing SIM files: {sim_files}")
            i = 0
            while i < len(sim_files):
                if not self.scanning:
                    break

                sim_file1 = sim_files[i]
                sim_file2 = sim_files[i + 1] if i + 1 < len(sim_files) else None

                if sim_file2 and self.are_sim_files_matching(sim_file1.name, sim_file2.name):
                    self.sim_file_manager.update_bat_file(bat_file1, com1, folder_path, sim_file1.name)
                    self.sim_file_manager.update_bat_file(bat_file2, com2, folder_path, sim_file2.name)
                    output_file1, output_file2 = self.sim_file_manager.run_bat_files(bat_file1, bat_file2)
                    i += 2
                else:
                    self.sim_file_manager.update_bat_file(bat_file1, com1, folder_path, sim_file1.name)
                    output_file1 = "log_batch_1.txt"
                    self.sim_file_manager.run_bat_file(bat_file1, output_file1)
                    output_file2 = None
                    i += 1

                vin_updated = False
                for record in data:
                    expected_prefix = f"{record['Year']} {record['Make']} {record['Model']} {record['Engine']}"
                    if all(part in sim_file1.name for part in expected_prefix.split()):
                        vin = record['VIN']
                        self.update_setting(vin, selected_folder)
                        vin_updated = True
                        break

                if not vin_updated:
                    logging.error(f"No matching record found for SIM file: {sim_file1.name}")

                time.sleep(3)

                if self.restart_device_var.get():
                    self.device_manager.restart_device()

                while not self.device_manager.check_device_connection():
                    if not self.scanning:
                        break
                    logging.info("Device not connected. Waiting for 10 seconds...")
                    time.sleep(10)

                self.device_manager.restart_app('com.innova.passthru')

                if not self.sim_file_manager.check_bat_file_output(output_file1, output_file2):
                    logging.error("Failed to connect with bat files after app restart.")
                    continue

                self.run_each_VIN(selected_folder)
                self.sim_file_manager.stop_running_processes()

            logging.info("Completed processing all SIM files.")

        threading.Thread(target=run_all).start()

    def run_each_VIN(self, selected_folder):
        global text
        max_retries = 2
        while True:
            retries = 0
            while retries < max_retries:
                try:
                    self.device_manager.restart_uiautomator2_server()
                    WebDriverWait(self.device_manager.driver, 50).until(
                        lambda driver: self.device_manager.check_device_connection())
                    self.check_memory_usage()
                    if not self.device_manager.driver or not self.device_manager.driver.session_id:
                        raise Exception("Appium driver is not properly initialized.")
                    self.find_VIN_mainscreen()
                    text = self.find_VIN_text()
                    logging.info(f"VIN Text: {text}")
                    if text and "NO VEHICLE INFORMATION" not in text:
                        self.write_VIN_to_txt(config.txt_path, text)

                        if self.all_var.get():
                            logging.info("Running all functions")
                            self.run_all_functions()
                            time.sleep(2)
                            logging.info("Running NWS live data")
                            self.run_nws_livedata()
                        else:
                            if self.obd2_10modes_var.get():
                                logging.info("Running OBD2 10 Modes functions")
                                self.run_obd2_10modes()

                            if self.obd2_livedata_var.get():
                                logging.info("Running OBD2 LiveData functions")
                                self.run_obd2_livedata()

                            if self.nws_dtcs_var.get():
                                logging.info("Running NWS DTCs functions")
                                self.run_nws_dtcs()

                            if self.nws_livedata_var.get():
                                logging.info("Running NWS live data")
                                self.run_nws_livedata()
                        # self.obd2_led_logic_var = tk.BooleanVar()

                        self.root.after(5000, self.stop_loading_animation)
                        return
                    else:
                        logging.warning("Do not connect with data sim files")
                        self.device_manager.restart_app("com.innova.passthru")
                        retries += 1
                except Exception:
                    logging.warning(f"Attempt {retries + 1} failed")
                    retries += 1
            if retries == max_retries:
                text = self.find_VIN_text()

                logging.warning("Failed to connect with data sim files after multiple attempts")
                self.write_VIN_to_txt(config.txt_path, text)
                break
            else:
                retries = 0

    def write_VIN_to_txt(self, file_path, text):
        with open(file_path, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([text])

    def find_VIN_text(self):
        start_time = time.time()
        vin_text = None
        xpaths = [
            "//android.webkit.WebView[@text='Ionic App']/android.view.View/android.view.View/android.view.View/android.view.View/android.view.View[2]/android.view.View/android.view.View[1]//android.view.View[@text]",
            "//android.webkit.WebView[@text='Ionic App']/android.view.View/android.view.View/android.view.View/android.view.View/android.view.View[2]/android.view.View//android.view.View[@text]",
            "//android.webkit.WebView[@text='Ionic App']/android.view.View/android.view.View/android.view.View/android.view.View/android.view.View[2]/android.view.View/android.view.View[1]"
        ]
        while time.time() - start_time < 60:
            for xpath in xpaths:
                try:
                    el = WebDriverWait(self.device_manager.driver, 10).until(
                        EC.presence_of_element_located((AppiumBy.XPATH, xpath)))
                    vin_text = el.get_attribute("text")
                    logging.info(f"Found VIN: {vin_text}")
                    if vin_text:
                        return vin_text
                except (TimeoutException, NoSuchElementException):
                    logging.warning(f"Element not found for XPath: {xpath}")
                except Exception:
                    logging.exception(f"Unexpected error with XPath {xpath}")
            time.sleep(1)
        logging.warning("Failed to find VIN within the time limit")
        return vin_text
    def find_VIN_mainscreen(self):
        def is_toyota_car():
            for _ in range(3):
                elements = self.device_manager.driver.find_elements(by=AppiumBy.XPATH, value="//android.view.View[@text]")
                for element in elements:
                    text = element.get_attribute("text")
                    if "Toyota" in text:
                        return True
                time.sleep(2)
            return False

        def click_button(xpath):
            try:
                element = self.device_manager.driver.find_element(by=AppiumBy.XPATH, value=xpath)
                element.click()
                return True
            except:
                return False

        if is_toyota_car():
            time.sleep(5)
            if click_button("//android.widget.Button[@text='TMMC, TMMK Product']"):
                time.sleep(5)
                click_button(
                    "//android.app.Dialog/android.view.View/android.view.View[2]/android.view.View/android.view.View[3]/android.view.View/android.view.View[1]/android.view.View")
                time.sleep(5)
                click_button("//android.app.Dialog/android.view.View/android.view.View[3]/android.view.View[4]")
            else:
                time.sleep(5)
                if click_button("//android.widget.Button[@text='w/ Smart Key']") or click_button(
                        "//android.widget.Button[@text='w/ ADK Package']"):
                    click_button("//android.widget.Button[@text='RADAR CRUISE']") or click_button(
                        "//android.widget.Button[@text='w/ EPB']")
                    click_button("//android.widget.Button[@text='Yes']")

    def check_and_wait_loading_xpath(self, xpath, wait_time=1.5):
        time.sleep(wait_time)
        try:
            self.device_manager.driver.find_element(by=AppiumBy.XPATH, value=xpath)
            logging.info(f"'{xpath}' appeared after {wait_time} seconds, waiting for it to disappear.")
            self.wait_for_element_disappearance(xpath)
        except:
            logging.info(f"'{xpath}' did not appear, assuming it already disappeared.")

    def wait_for_element_disappearance(self, xpath, timeout=30, interval=0.5):
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                self.device_manager.driver.find_element(by=AppiumBy.XPATH, value=xpath)
                time.sleep(interval)
            except:
                end_time = time.time()
                logging.info(f"Element {xpath} disappeared after {end_time - start_time:.2f} seconds.")
                return True
        logging.warning(f"Timeout: Element {xpath} did not disappear within {timeout} seconds.")
        return False
    def check_memory_usage(self):
        process = psutil.Process(os.getpid())
        memory_usage = process.memory_info().rss / (1024 * 1024)
        logging.info(f"Memory usage: {memory_usage} MB")
        if memory_usage > 1000:
            logging.warning("High memory usage detected. Restarting driver.")
            self.device_manager.restart_uiautomator2_server()

    def load_data_excel(self, sheet_name):
        try:
            with open(config.setting_path, 'r') as json_file:
                settings = json.load(json_file)
            path_file = config.database_dir / settings['test_document']
            df = pd.read_excel(path_file, sheet_name=sheet_name, usecols=['Year', 'Make', 'Model', 'Engine', 'VIN'],
                               na_filter=False)
            return df.to_dict('records')
        except FileNotFoundError:
            logging.exception(f"File not found: {path_file}")
        except Exception:
            logging.exception(f"Error loading data from Excel sheet {sheet_name}")
        return []

    def remove_duplicates(self, data):
        unique_data = {frozenset(item.items()): item for item in data}
        return list(unique_data.values())

    def update_setting(self, vin, sheet_name):
        try:
            with open(config.setting_path, 'r') as json_file:
                settings = json.load(json_file)
            settings.update({'VIN': vin, 'sheet_name': sheet_name})
            with open(config.setting_path, 'w') as json_file:
                json.dump(settings, json_file, indent=4)
            logging.info(f"Updated setting.json with VIN: {vin} and sheet_name: {sheet_name}")
        except (FileNotFoundError, PermissionError):
            logging.exception(f"Error updating setting.json")
        except Exception:
            logging.exception(f"Unexpected error updating setting.json")

    def are_sim_files_matching(self, file1, file2):
        return file1.split('_')[0] == file2.split('_')[0]

    def run_all_functions(self):
        try:
            logging.info("Running all functions sequentially...")
            subprocess.run(["python", config.all_functions], check=True)
            logging.info("Success: All functions execution completed.")
        except subprocess.CalledProcessError:
            logging.exception("Error occurred while running all functions")
        except Exception:
            logging.exception("Unexpected error occurred")

    def run_nws_livedata(self):
        try:
            logging.info("Running NWS live data sequentially...")
            subprocess.run(["python", config.nws_live_data_functions], check=True)
            logging.info("Success: NWS LiveData execution completed.")
        except subprocess.CalledProcessError:
            logging.exception("Error occurred while running NWS live data")
        except Exception:
            logging.exception("Unexpected error occurred")
    def run_obd2_10modes(self):
        try:
            logging.info("Running OBD2 10 Modes...")
            subprocess.run(["python", config.obd2_10modes], check=True)
            logging.info("Success:  OBD2 10 Modes execution completed.")
        except subprocess.CalledProcessError:
            logging.exception("Error occurred while running OBD2 10 Modes")
        except Exception:
            logging.exception("Unexpected error occurred")
    def run_obd2_livedata(self):
        try:
            logging.info("Running OBD2 LiveData...")
            subprocess.run(["python", config.obd2_livedata], check=True)
            logging.info("Success: OBD2 LiveData execution completed.")
        except subprocess.CalledProcessError:
            logging.exception("Error occurred while running OBD2 LiveData")
        except Exception:
            logging.exception("Unexpected error occurred")

    def run_nws_dtcs(self):
        try:
            logging.info("Running NWS DTCs sequentially...")
            subprocess.run(["python", config.nws_dtcs], check=True)
            logging.info("Success: NWS DTCs execution completed.")
        except subprocess.CalledProcessError:
            logging.exception("Error occurred while running NWS DTCs ")
        except Exception:
            logging.exception("Unexpected error occurred")
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
