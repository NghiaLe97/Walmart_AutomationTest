import json
import logging
import subprocess
import sys
import unittest
from appium.options.common.base import AppiumOptions
from appium import webdriver  # Corrected import
import Actions
from time import sleep
import mil_color
from pathlib import Path


def get_base_dir():
    return getattr(sys, 'frozen', False) and sys._MEIPASS or Path(__file__).resolve().parent


base_dir = get_base_dir()
simfile_path = base_dir.parent / "Sim 01 41"
database_dir = base_dir.parent / "database"
document_path = database_dir / "database.json"


def restart_appium_server():
    subprocess.run(["taskkill", "/IM", "node.exe", "/F"], shell=True)  # Windows
    subprocess.Popen(["appium", "--log-level", "debug"], shell=True)


def restart_uiautomator2(device_id):
    subprocess.run(f"adb -s {device_id} shell am force-stop io.appium.uiautomator2.server", shell=True)
    subprocess.run(f"adb -s {device_id} shell am start -n io.appium.uiautomator2.server/.MainActivity", shell=True)


def free_port(device_id, port=8206):
    subprocess.run(f"adb -s {device_id} forward --remove tcp:{port}", shell=True)


def is_device_connected(device_id):
    result = subprocess.run(f"adb devices", shell=True, capture_output=True, text=True)
    return device_id in result.stdout


class TestApp(unittest.TestCase):
    config = None
    server_url = "http://localhost:4723"
    device_id = "172.32.0.152:38759"  # ID thiết bị của bạn

    def setUp(self):
        # Đảm bảo thiết bị được kết nối
        if not is_device_connected(self.device_id):
            logging.warning("Device not connected. Attempting to reconnect...")
            subprocess.run(f"adb reconnect {self.device_id}", shell=True)

        free_port(self.device_id)
        restart_uiautomator2(self.device_id)

        options = AppiumOptions()
        options.load_capabilities({
            "platformName": "android",
            "appium:automationName": "uiautomator2",
            "appium:noReset": True,
            "newCommandTimeout": 3600,
            "uiautomator2ServerLaunchTimeout": 120000,
            "adbExecTimeout": 60000
        })
        try:
            self.driver = webdriver.Remote(self.server_url, options=options)  # Corrected usage
        except Exception as e:
            restart_appium_server()
            self.fail(f"Failed to initialize Appium session: {e}")

    def update_setting(self, function, value):
        try:
            with open(document_path, 'r') as json_file:
                settings = json.load(json_file)
            settings.update({f'{function}': value})
            with open(document_path, 'w') as json_file:
                json.dump(settings, json_file, indent=4)
        except (FileNotFoundError, PermissionError):
            logging.exception(f"Error updating setting.json")
        except Exception:
            logging.exception(f"Unexpected error updating setting.json")

    def test_main(self):
        Auto(self)

    def tearDown(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception as e:
                logging.warning(f"Error during driver quit: {e}")


def Auto(self, retries=3):
    for attempt in range(retries):
        try:
            success, message = Actions.check_communication_error(self)
            if not success:
                print(f"Test failed: {message}")

            sleep(5)

            success, message = Actions.click_Setting(self)
            if not success:
                print(f"Test failed: {message}")

            sleep(5)

            success, message = Actions.click_IMProgramLocation(self)
            if not success:
                print(f"Test failed: {message}")

            sleep(5)

            success, message = Actions.click_show_listLocation(self)
            if not success:
                print(f"Test failed: {message}")

            sleep(5)

            success, message = Actions.click_Location(self)
            if not success:
                print(f"Test failed: {message}")

            sleep(5)

            success, message = Actions.click_back(self)
            if not success:
                print(f"Test failed: {message}")

            sleep(5)
            success, message = Actions.check_communication_error(self)
            if not success:
                print(f"Test failed: {message}")

            sleep(5)

            success, message = Actions.click_OBD2Diagnostics(self)
            if not success:
                print(f"Test failed: {message}")
            sleep(15)

            result_mil = mil_color.check_led(self)
            print(f"MIL color: {result_mil}")
            self.update_setting("LED", result_mil)

            break
        except Exception as e:
            logging.error(f"Error occurred: {e}. Retrying ({attempt + 1}/{retries})...")
            if self.driver:
                self.driver.quit()  # Đóng session hiện tại
            restart_appium_server()  # Khởi động lại Appium server
            self.setUp()  # Khởi tạo lại Appium session
    else:
        self.fail("Failed to complete Auto process after retries.")


if __name__ == "__main__":
    unittest.main()