import logging
import subprocess
import time
from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.common.appiumby import AppiumBy
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

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
