import sys

from PIL import Image, ImageStat
from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.common.appiumby import AppiumBy
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import os
import re
import json
from pathlib import Path


def get_base_dir():
    return getattr(sys, 'frozen', False) and sys._MEIPASS or Path(__file__).resolve().parent


base_dir = get_base_dir()
simfile_path = base_dir.parent / "Sim 01 41"
database_dir = base_dir.parent / "database"
document_path = database_dir / "database.json"
# document_path = "database/database.json"
screenshot_dir_parent = base_dir / "Screenshot"
screenshot_dir = screenshot_dir_parent / "MIL"

case_colors = {
    "Case 1 (Red)": (233, 35, 35),
    "Case 2 (Yellow)": (221, 173, 83),
    "Case 3 (Green)": (21, 166, 66),
    "Non": (50, 67, 59)  # Màu trung tính cho trường hợp không sáng
}
color_tolerance = 30  # Độ lệch màu cho phép


def check_led(self):
    try:
        with open(document_path, 'r') as json_file:
            settings = json.load(json_file)
        case_value = settings.get("Case", "UnknownCase").replace(" ", "_").replace(".", "_")

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        element = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((AppiumBy.XPATH,
                                            '//android.webkit.WebView[@text="Ionic '
                                            'App"]/android.view.View/android.view.View/android.view.View/android.view'
                                            '.View/android.view.View['
                                            '2]/android.view.View/android.view.View/android.view.View/android.view'
                                            '.View/android.view.View/android.view.View/android.view.View/android.view'
                                            '.View[3]/android.view.View'))
        )

        bounds = element.get_attribute("bounds")
        matches = re.findall(r"\d+", bounds)
        x1, y1, x2, y2 = map(int, matches)

        full_screenshot_path = os.path.join(screenshot_dir, f"MIL_{case_value}_{timestamp}.png")
        self.driver.save_screenshot(full_screenshot_path)

        img = Image.open(full_screenshot_path)
        cropped_img = img.crop((2243, y1, 2483, y2))
        cropped_screenshot_path = os.path.join(screenshot_dir, f"MIL_Cropped_{case_value}_{timestamp}.png")
        cropped_img.save(cropped_screenshot_path)

        # Xác định vị trí tương đối của ba hình tròn
        circle_width = (2473 - 2243) // 3
        circles = [
            cropped_img.crop((circle_width * i, 0, circle_width * (i + 1), cropped_img.height))
            for i in range(3)
        ]

        brightness_values = []
        for idx, circle in enumerate(circles):
            # Tính độ sáng của hình tròn
            stat = ImageStat.Stat(circle)
            brightness = sum(stat.mean[:3]) / 3
            brightness_values.append((idx + 1, brightness))
            # print(f"Hình tròn {idx + 1} - Độ sáng trung bình: {brightness}")

        # Sắp xếp độ sáng từ cao đến thấp để xác định hình tròn nào sáng nhất
        brightness_values.sort(key=lambda x: x[1], reverse=True)

        # Kiểm tra điều kiện để xác định màu cuối cùng
        if brightness_values[0][1] > brightness_values[1][1] + 2:  # Ngưỡng để xác định chênh lệch độ sáng
            if brightness_values[0][0] == 1:
                return "Red"
            elif brightness_values[0][0] == 2:
                return "Yellow"
            elif brightness_values[0][0] == 3:
                return "Green"
        else:
            return "Fail"
    except Exception as e:
        print("Không tìm thấy phần tử:", e)
