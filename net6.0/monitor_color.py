import sys

from PIL import Image
from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.common.appiumby import AppiumBy
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, WebDriverException
import collections
import re
import webcolors
import os
import json
from datetime import datetime
from pathlib import Path


def get_base_dir():
    return getattr(sys, 'frozen', False) and sys._MEIPASS or Path(__file__).resolve().parent


base_dir = get_base_dir()
simfile_path = base_dir.parent / "Sim 01 41"
database_dir = base_dir.parent / "database"
document_path = database_dir / "database.json"
screenshot_dir_parent = base_dir / "Screenshot"
screenshot_dir = screenshot_dir_parent / "MONITOR"

# Định nghĩa màu CSS3
css3_colors = {
    "red": (255, 0, 0),
    "gray": (128, 128, 128),
    "green": (0, 128, 0),
}


def closest_color(requested_color, threshold=1000):
    min_distance = None
    closest_name = None
    for name, rgb in css3_colors.items():
        distance = sum((a - b) ** 2 for a, b in zip(rgb, requested_color))
        if (min_distance is None or distance < min_distance) and distance < threshold:
            min_distance = distance
            closest_name = name
    return closest_name if closest_name else "không xác định"


def get_color_name(rgb_color):
    try:
        return webcolors.rgb_to_name(rgb_color)
    except ValueError:
        return closest_color(rgb_color)


def determine_dominant_color(color_counts, threshold_ratio=0.5):
    green_count = color_counts.get("green", 0)
    red_count = color_counts.get("red", 0)
    gray_count = color_counts.get("gray", 0)

    if green_count > 0 and red_count == 0:
        if gray_count >= green_count * threshold_ratio:
            return "Green Gray"
        else:
            return "Green"
    elif red_count > 0 and green_count == 0:
        if gray_count >= red_count * threshold_ratio:
            return "Red Gray"
        else:
            return "Red"
    return "Fail"


def check_element_colors(server_url='http://localhost:4723', device_name="AndroidDevice"):
    try:
        with open(document_path, 'r') as json_file:
            settings = json.load(json_file)
        case_value = settings.get("Case", "UnknownCase").replace(" ", "_").replace(".", "_")
    except Exception as e:
        print("File database.json:", e)
        case_value = "UnknownCase"

    # Tạo timestamp cho tên file
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    screenshot_path = os.path.join(screenshot_dir, f"MONITOR_{case_value}_{timestamp}.png")

    options = UiAutomator2Options()
    options.platform_name = "Android"
    options.device_name = device_name
    options.new_command_timeout = 600
    options.skip_server_installation = True
    options.uiautomator2_server_launch_timeout = 120000

    try:
        driver = webdriver.Remote(command_executor=server_url, options=options)
    except WebDriverException as e:
        print("Session:", e)
        return "Session Init Failed"

    color_results = []

    try:
        xpath_template = '//android.webkit.WebView[@text="Ionic App"]/android.view.View/android.view.View/android.view.View/android.view.View/android.view.View[2]/android.view.View/android.view.View/android.view.View/android.view.View/android.view.View/android.view.View/android.view.View/android.view.View[{}]/android.view.View/android.view.View[2]/android.view.View[{}]'

        khung_max = None
        for i in range(8, 1, -1):
            xpath = xpath_template.format(i, 1)
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((AppiumBy.XPATH, xpath)))
                khung_max = i
                break
            except Exception:
                continue

        if not khung_max:
            return "Failed"

        detail_max = None
        for j in range(11, 1, -1):
            xpath = xpath_template.format(khung_max, j)
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((AppiumBy.XPATH, xpath)))
                detail_max = j
                break
            except Exception:
                continue

        if not detail_max:
            return "Failed"

        for j in range(detail_max - 1, max(0, detail_max - 6), -1):
            xpath = xpath_template.format(khung_max, j)
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((AppiumBy.XPATH, xpath))
                )
                bounds = element.get_attribute("bounds")
            except Exception:
                continue

            matches = re.findall(r"\d+", bounds)
            x1, y1, x2, y2 = map(int, matches)

            driver.save_screenshot(screenshot_path)  # Lưu ảnh toàn màn hình
            img = Image.open(screenshot_path)

            color_counts = collections.Counter()
            for x in range(x1, x2):
                for y in range(y1, y2):
                    pixel_color = img.getpixel((x, y))
                    if len(pixel_color) == 4:
                        pixel_color = pixel_color[:3]
                    color_name = get_color_name(pixel_color)
                    color_counts[color_name] += 1

            dominant_color = determine_dominant_color(color_counts)
            color_results.append(dominant_color)

    finally:
        try:
            driver.quit()
        except WebDriverException:
            print("Error")

    if len(color_results) > 0 and all(result == color_results[0] for result in color_results):
        return color_results[0]
    else:
        return "Fail"
