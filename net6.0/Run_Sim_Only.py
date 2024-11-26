import logging
import time
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import subprocess
import serial.tools.list_ports

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class Config:
    def __init__(self):
        self.base_dir = Path(__file__).resolve().parent
        self.simfile_path = self.base_dir.parent / "Divide Make"
        self.bat_file_1 = self.base_dir / "SimulationWithShowData.bat"
        self.bat_file_2 = self.base_dir / "SimulationWithShowData2.bat"
        self.ensure_directories_exist()

    def ensure_directories_exist(self):
        directories = [self.simfile_path]
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            logging.info(f"Directory exists or created: {directory}")


config = Config()


class SimFileManager:
    def __init__(self, sim_files_path):
        self.sim_files_path = sim_files_path
        self.processes = []

    def update_bat_file(self, bat_file, com_port, folder_path, sim_file):
        try:
            with open(bat_file, 'r') as file:
                content = file.read()
            parts = content.split('"')
            parts[1] = com_port
            parts[3] = str(Path(folder_path) / sim_file)
            new_content = '"'.join(parts)
            with open(bat_file, 'w') as file:
                file.write(new_content)
            logging.info(f"Updated {bat_file} with COM port {com_port} and SIM file {sim_file}")
        except Exception as e:
            logging.error(f"An error occurred while updating {bat_file}: {e}")

    def run_bat_file(self, bat_file):
        command = f'start cmd.exe /c "{bat_file}"'
        process = subprocess.Popen(
            command,
            shell=True
        )
        self.processes.append(process)  # Save process into the list

    def run_bat_files(self, bat_file1, bat_file2=None):
        self.stop_running_processes()
        time.sleep(3)

        self.run_bat_file(bat_file1)
        if bat_file2:
            self.run_bat_file(bat_file2)
        logging.info(f"Running batch files: {bat_file1}, {bat_file2}")

    def stop_running_processes(self):
        subprocess.call(["taskkill", "/F", "/IM", "SimulatorTest.exe"], stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL)
        subprocess.call(["taskkill", "/F", "/IM", "cmd.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Stopped running processes: SimulatorTest.exe, cmd.exe")


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automation SIM Runner")
        self.root.geometry("400x250")
        self.sim_file_manager = SimFileManager(config.simfile_path)
        self.create_widgets()
        self.update_folder_combobox()

    def create_widgets(self):
        frame = ttk.Frame(self.root)
        frame.pack(padx=20, pady=20)
        self.combobox_hid1 = self.create_combobox(frame, "HID 1:")
        self.combobox_hid2 = self.create_combobox(frame, "HID 2:")
        self.folder_combobox = self.create_combobox(frame, "Make:")
        self.scan_folder_button = ttk.Button(frame, text="Scan Folder", command=self.scan_folder)
        self.scan_folder_button.pack(pady=5)
        self.next_button = ttk.Button(frame, text="NEXT", command=self.process_next_sim)
        self.next_button.pack(pady=5)
        self.sim_files = []
        self.current_index = -1  # To track which SIM file is currently being processed

    def create_combobox(self, parent, label_text):
        label = ttk.Label(parent, text=label_text)
        label.pack()
        combobox = ttk.Combobox(parent, state="readonly")
        combobox.pack()

        # Thêm sự kiện để làm mới khi click vào Combobox
        combobox.bind("<Button-1>", self.refresh_com_ports)
        return combobox

    def update_folder_combobox(self):
        try:
            folders = [f for f in config.simfile_path.iterdir() if f.is_dir()]
            self.folder_combobox['values'] = [f.name for f in folders]
            logging.info(f"Available folders: {[f.name for f in folders]}")
        except FileNotFoundError:
            logging.exception("Error accessing SIM files directory")
            messagebox.showerror("Error", "Could not access SIM files directory.")

    def refresh_com_ports(self, event=None):
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.combobox_hid1['values'] = port_list
        self.combobox_hid2['values'] = port_list
        logging.info(f"Available COM ports: {port_list}")

    def scan_folder(self):
        selected_folder = self.folder_combobox.get()
        if not selected_folder:
            messagebox.showwarning("Warning", "Please select a folder.")
            logging.warning("No folder selected.")
            return

        folder_path = config.simfile_path / selected_folder
        # Exclude files that end with '.correct.sim'
        self.sim_files = [f for f in folder_path.iterdir() if
                          f.suffix == '.sim' and not f.name.endswith('.correct.sim')]
        logging.info(f"Scanning folder: {selected_folder} with SIM files: {self.sim_files}")

        if not self.sim_files:
            messagebox.showwarning("Warning", "No SIM files found in the selected folder.")
            logging.warning("No SIM files found in the selected folder.")
            return

        self.current_index = -1
        self.process_next_sim()

    def process_next_sim(self):
        self.current_index += 1

        if self.current_index >= len(self.sim_files):
            messagebox.showinfo("Info", "All SIM files processed.")
            logging.info("All SIM files processed.")
            return

        sim_file1 = self.sim_files[self.current_index]
        sim_file2 = self.sim_files[self.current_index + 1] if self.current_index + 1 < len(self.sim_files) else None

        com1 = self.combobox_hid1.get()
        com2 = self.combobox_hid2.get()

        if not com1 or not com2:
            messagebox.showwarning("Warning", "Please select both COM ports.")
            logging.warning("COM ports not selected.")
            return

        if sim_file2 and self.are_sim_files_matching(sim_file1.name, sim_file2.name):
            # When there are 2 matching SIM files
            self.sim_file_manager.update_bat_file(config.bat_file_1, com1, sim_file1.parent, sim_file1.name)
            self.sim_file_manager.update_bat_file(config.bat_file_2, com2, sim_file2.parent, sim_file2.name)
            self.sim_file_manager.run_bat_files(config.bat_file_1, config.bat_file_2)
            self.current_index += 1  # Skip the next file since it's already processed
        else:
            # When there is only 1 SIM file
            self.sim_file_manager.update_bat_file(config.bat_file_1, com1, sim_file1.parent, sim_file1.name)
            self.sim_file_manager.run_bat_files(config.bat_file_1)

    def are_sim_files_matching(self, file1, file2):
        return file1.split('_')[0] == file2.split('_')[0]


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
