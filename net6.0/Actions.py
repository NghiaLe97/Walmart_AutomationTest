import sys

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from appium.webdriver.common.appiumby import AppiumBy
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep
import json
from pathlib import Path


def get_base_dir():
    return getattr(sys, 'frozen', False) and sys._MEIPASS or Path(__file__).resolve().parent


base_dir = get_base_dir()
simfile_path = base_dir.parent / "Sim 01 41"
database_dir = base_dir.parent / "database"
document_path = database_dir / "database.json"


def swipe_up_eight_times(self):
    try:
        # Swipe from (550, 550) to (550, 350) 8 times
        for i in range(5):
            self.driver.swipe(550, 550, 550, 350)
            sleep(0.5)
        return True, "Swiping completed successfully."

    except Exception as e:
        print(f"An error occurred while swiping: {e}")
        return False, f"Error while swiping: {e}"


def swipe_up_location(self):
    try:
        for i in range(8):
            self.driver.swipe(2000, 1123, 2000, 810)
            sleep(0.5)  # Add a short delay between swipes, adjust as needed
        return True, "Swiping completed successfully."
    except Exception as e:
        print(f"An error occurred while swiping: {e}")
        return False, f"Error while swiping: {e}"


def swipe_up_long(self):
    try:
        for i in range(8):
            self.driver.swipe(1300, 1200, 1300, 400)
            sleep(0.5)  # Add a short delay between swipes, adjust as needed
        return True, "Swiping completed successfully."
    except Exception as e:
        print(f"An error occurred while swiping: {e}")
        return False, f"Error while swiping: {e}"


def click_back(self):
    try:
        click_Back = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//android.widget.Image[@text="arrow round back"]'))
        )
        click_Back.click()
        print("Successfully click Back button")
        return True, None

    except TimeoutException:
        message = "Error when click Back button"
        print(message)
        return False, message


def click_home(self):
    try:
        click_Back = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//android.webkit.WebView[@text="Ionic '
                           'App"]/android.view.View/android.view.View/android.view.View/android.view.View/android'
                           '.view.View[1]/android.view.View/android.view.View[2]/android.view.View['
                           '3]/android.view.View/android.widget.Button'))
        )
        click_Back.click()
        print("Successfully click Home button")
        return True, None

    except TimeoutException:
        message = "Error when click Hone button"
        print(message)
        return False, message


def check_communication_error(self):
    try:
        # Continuously check for "Communication Error" and click "Cancel" until error disappears
        while True:
            try:
                # Check if "Communication Error" is displayed
                check_Communication_error = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//android.widget.TextView[@text="Communication Error"] '
                                                              '| //android.view.View[@text="Communication Error"]'))
                )

                if check_Communication_error:
                    # Locate and click the "Cancel" button if error is detected
                    cancel_button = self.driver.find_element(AppiumBy.XPATH, '//android.widget.Button[@text="Cancel"]')
                    cancel_button.click()
                    print("Detected Communication Error - clicked Cancel to retry")

                    # Brief pause before rechecking to allow for UI update
                    sleep(2)

            except TimeoutException:
                # No "Communication Error" found, meaning error has disappeared, so we exit
                print("Communication Error resolved. Exiting check.")
                return True, None

    except Exception as e:
        # Handle unexpected issues and return a specific error message
        message = f"Error when loading SIM file: {e}"
        print(message)
        return False, message


def check_detect_vehicle(self):
    try:
        # Continuously check for "Communication Error" and click "Cancel" until error disappears
        while True:
            try:
                # Check if "Communication Error" is displayed
                check_Communication_error = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '//android.widget.TextView[@text="Cannot detect Vehicle Information Number (VIN).\nDo you want to go to Vehicle Selection?"]'))
                )

                if check_Communication_error:
                    # Locate and click the "Cancel" button if error is detected
                    cancel_button = self.driver.find_element(AppiumBy.XPATH, '//android.widget.Button[@text="No"]')
                    cancel_button.click()
                    print("Detected Communication Error - clicked Cancel to retry")

                    # Brief pause before rechecking to allow for UI update
                    sleep(2)

            except TimeoutException:
                # No "Communication Error" found, meaning error has disappeared, so we exit
                print("Communication Error resolved. Exiting check.")
                return True, None

    except Exception as e:
        # Handle unexpected issues and return a specific error message
        message = f"Error when loading SIM file: {e}"
        print(message)
        return False, message


def click_OBD2Diagnostics(self):
    try:
        click_OBD2_Diagnostics = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//android.widget.Button[@text="information circle outline OBD2 Diagnostics"]'))
        )
        click_OBD2_Diagnostics.click()
        print("Successfully click OBD2 Diagnostics")
        return True, None

    except TimeoutException:
        message = "Error when click OBD2 Diagnostics"
        print(message)
        return False, message


def check_vehicle_selection(self):
    try:
        # Load data from database.json file
        with open(document_path, "r") as file:
            data = json.load(file)
        years = data.get("Years", "")

        # Define expected condition based on "Years" keyword
        requires_vehicle_selection = years.startswith("Before") or years.startswith("After")

        try:
            # Wait for "Vehicle Selection" to be present if expected
            vehicle_selection = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//android.widget.TextView[@text="Vehicle Selection"]'))
            )

            if requires_vehicle_selection:
                print("'Vehicle Selection' appeared as expected for years:", years)
            else:
                print("Unexpected 'Vehicle Selection' found for years:", years)
                return False, f"Bug: 'Vehicle Selection' appeared unexpectedly for years: {years}"

            # Click the "Select New Vehicle" button
            select_new_vehicle_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//android.widget.Button[@text="Select New Vehicle"]'))
            )
            select_new_vehicle_button.click()
            print("Clicked 'Select New Vehicle' button")
            sleep(3)

            # Proceed to select year and make
            check_year_and_make(self)
            sleep(1)
            return True, "Vehicle selection successful"

        except TimeoutException:
            # Handle cases where "Vehicle Selection" is missing
            if requires_vehicle_selection:
                print(f"Bug: 'Vehicle Selection' not found as expected for years: {years}")
                return False, f"Bug: 'Vehicle Selection' missing for years: {years}"
            else:
                print("No 'Vehicle Selection' found, which is correct for years:", years)
                return True, "Vehicle selection not needed; skipped action"

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return False, f"Unexpected error: {e}"


def check_toyota_selection(self):
    try:
        # Load data from database.json file
        with open(document_path, "r") as file:
            data = json.load(file)
        years = data.get("Years", "")

        # Define expected condition based on "Years" keyword
        requires_vehicle_selection = years.startswith("Toyota")

        try:
            # Wait for "Vehicle Selection" to be present if expected
            vehicle_selection = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//android.widget.TextView[@text="Select Option 1"]'))
            )

            if requires_vehicle_selection:
                print("'Vehicle Selection' appeared as expected for years:", years)
            else:
                print("Unexpected 'Vehicle Selection' found for years:", years)
                return False, f"Bug: 'Vehicle Selection' appeared unexpectedly for years: {years}"

            # Click the "Select New Vehicle" button
            select_new_vehicle_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//android.widget.Button[@text="w/ Smart Key"]'))
            )
            select_new_vehicle_button.click()
            print("Clicked 'Select New Vehicle' button")
            sleep(3)
            click_yesYMME(self)
            return True, "Vehicle selection successful"

        except TimeoutException:
            # Handle cases where "Vehicle Selection" is missing
            if requires_vehicle_selection:
                print(f"Bug: 'Vehicle Selection' not found as expected for years: {years}")
                return False, f"Bug: 'Vehicle Selection' missing for years: {years}"
            else:
                print("No 'Vehicle Selection' found, which is correct for years:", years)
                return True, "Vehicle selection not needed; skipped action"

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return False, f"Unexpected error: {e}"


def check_year_and_make(self):
    # Load data from database.json file
    try:
        with open(document_path, "r") as file:
            data = json.load(file)
        print("Data successfully loaded from database.json:", data)
    except Exception as e:
        print(f"Error loading database.json file: {e}")
        return False, f"Error loading database.json file: {e}"

    # Retrieve "Years" and "Make" values from the data
    years = data.get("Years", "")
    make = data.get("Make", "")

    try:
        # Select year based on the "Years" value
        if years.startswith("After"):
            year_button = self.driver.find_element(By.XPATH, '//android.widget.Button[@text="2008"]')
            year_button.click()
            print(f"Clicked '2008' button for year {years}")
        elif years.startswith("Before"):
            year_button = self.driver.find_element(By.XPATH, '//android.widget.Button[@text="2005"]')
            year_button.click()
            print(f"Clicked '2005' button for year {years}")
        else:
            print("Invalid year value:", years)
            return False, "Invalid year value"

        sleep(2)

        # Select "Make" based on specific conditions or use default
        make_button_xpath = {
            "JaguarLandrover": '//android.widget.Button[@text="Land Rover"]',
            "GM": '//android.widget.Button[@text="GMC"]',
            "Mercedes": '//android.widget.Button[@text="Mercedes-Benz"]'
        }.get(make, f'//android.widget.Button[@text="{make}"]')

        try:
            make_button = self.driver.find_element(By.XPATH, make_button_xpath)
            make_button.click()
            print(f"Clicked '{make}' button for Make '{make}'")
        except NoSuchElementException:
            print(f"Make button '{make}' not found in UI.")
            return False, f"Make button '{make}' not found in UI."

        sleep(2)

        # Execute the steps to select Model, Trim, Option, Engine, and click_yesYMME
        steps = [
            ("Model", select_Model),
            ("Trim", select_Trim),
            ("Option", select_Option),
            ("Engine", select_Engine),
            ("YMME Confirmation", click_yesYMME)
        ]

        for step_name, step_function in steps:
            try:
                success, message = step_function(self)
                if success:
                    print(f"Successfully completed {step_name} selection.")
                else:
                    print(f"Skipped {step_name}: {message}")
            except Exception as e:
                print(f"{step_name} not available or an error occurred: {e}")
                continue  # Skip to the next step if any step fails or is missing

            sleep(1)

        return True, "Successfully completed year and vehicle type selection"

    except NoSuchElementException as e:
        print(f"UI element not found: {e}")
        return False, f"Element not found: {e}"

    except TimeoutException as e:
        print(f"Timeout error while locating UI element: {e}")
        return False, f"Timeout error: {e}"

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return False, f"Unexpected error: {e}"


def select_Model(self):
    try:
        select_model = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 '//android.app.Dialog/android.view.View/android.view.View[2]/android.view.View/android.view.View/android.view.View[3]/android.view.View[1]/android.view.View[1]/android.view.View'))
        )
        select_model.click()
        print("Successfully select Model")
        return True, None

    except TimeoutException:
        message = "Error when select Model"
        print(message)
        return False, message


def select_Trim(self):
    try:
        select_trim = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 '//android.app.Dialog/android.view.View/android.view.View[2]/android.view.View/android.view.View/android.view.View[3]/android.view.View/android.view.View[1]/android.view.View'))
        )
        select_trim.click()
        # print("Successfully select Trim")
        return True, None

    except TimeoutException:
        pass


def select_Option(self):
    try:
        select_option = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 '//android.app.Dialog/android.view.View/android.view.View[2]/android.view.View/android.view.View/android.view.View[3]/android.view.View/android.view.View[1]/android.view.View'))
        )
        select_option.click()
        # print("Successfully select Option")
        return True, None

    except TimeoutException:
        pass


def select_Engine(self):
    try:
        select_engine = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 '//android.app.Dialog/android.view.View/android.view.View[2]/android.view.View/android.view.View/android.view.View[3]/android.view.View/android.view.View[1]/android.view.View'))
        )
        select_engine.click()
        # print("Successfully select Engine")
        return True, None

    except TimeoutException:
        pass


def click_yesYMME(self):
    try:
        click_YesYMME = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//android.widget.Button[@text="Yes"]'))
        )
        click_YesYMME.click()
        print("Successfully click Yes button")
        return True, None

    except TimeoutException:
        message = "Error when click Yes button"
        print(message)
        return False, message


def check_DTC(self):
    try:
        # Check for specific DTC statuses and store found results
        confirmed_dtc = self.driver.find_elements(By.XPATH, '//android.widget.TextView[@text="Confirmed/MIL DTC"]')
        pending_dtc = self.driver.find_elements(By.XPATH, '//android.widget.TextView[@text="Pending DTC"]')
        permanent_dtc = self.driver.find_elements(By.XPATH, '//android.widget.TextView[@text="Permanent DTC"]')

        # Create a list to store statuses
        statuses = []
        if confirmed_dtc:
            statuses.append("Stored")
        if pending_dtc:
            statuses.append("Pending")
        if permanent_dtc:
            statuses.append("Permanent")

        # Return list of statuses if found, or "No DTC" if none
        if statuses:
            message = f"{'/'.join(statuses)}"
            print(f"DTC statuses found: {message}")
            return message
        else:
            message = "No"
            print(message)
            return message

    except Exception as e:
        # Return error message in case of an exception
        error_message = f"Error occurred: {e}"
        print(error_message)
        return ["Error", error_message]


def check_FF(self):
    try:
        # Check if the "Freeze Frame" button is present
        freeze_frame_button = WebDriverWait(self.driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//android.widget.Button[@text="Freeze Frame"]'))
        )

        if freeze_frame_button:
            print("Freeze Frame button found.")
            return "Yes"  # Return "Yes" if Freeze Frame button is found

    except TimeoutException:
        # If "Freeze Frame" button is not found, check for the DTC message
        try:
            dtc_message = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH,
                                                '//android.view.View[@text="No Powertrain DTCs or Freeze Frame Data is presently stored in the vehicle\'s computer."]'))
            )

            if dtc_message:
                print("No Powertrain DTCs or Freeze Frame Data is presently stored in the vehicle.")
                return "No"  # Return "No" if DTC message is found

        except TimeoutException:
            # If neither the Freeze Frame button nor the DTC message is found, return "No"
            print("Freeze Frame button and DTC message not found.")
            return "No"

    except Exception as e:
        # Handle any unexpected errors by returning "No"
        print(f"Error: {e}")
        return "No"


def check_MIL(self):
    try:
        mil_off = self.driver.find_elements(By.XPATH, '//android.widget.TextView[@text="MIL OFF"]')
        if mil_off:
            message = "OFF"
            # print(F"MIL status1: {message}")
            return message
    except NoSuchElementException:
        pass

    try:
        mil_on = self.driver.find_elements(By.XPATH, '//android.widget.TextView[@text="MIL ON"]')
        if mil_on:
            message = "ON"
            # print(F"MIL status1: {message}")
            return message
    except NoSuchElementException:
        pass

    print("MIL status not found")
    return "Failed"


# Setting
def click_Setting(self):
    try:
        click_setting = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//android.widget.Button[@text="information circle outline Settings"]'))
        )
        click_setting.click()
        print("Successfully click Setting")
        return True, None

    except TimeoutException:
        message = "Error when click Setting"
        print(message)
        return False, message


# Setting
def click_IMProgramLocation(self):
    try:
        swipe_up_eight_times(self)
        click_IM_ProgramLocation = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//android.widget.TextView[@text="I/M Program Location"]'))
        )
        click_IM_ProgramLocation.click()
        print("Successfully click I/M Program Location")
        return True, None

    except TimeoutException:
        message = "Error when click I/M Program Location"
        print(message)
        return False, message


# Setting
def click_show_listLocation(self):
    try:
        click_show_list_Location = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//android.webkit.WebView[@text="Ionic '
                           'App"]/android.view.View/android.view.View/android.view.View/android.view.View/android'
                           '.view.View[2]/android.view.View/android.view.View[2]/android.view.View['
                           '3]/android.view.View/android.widget.Image'))
        )
        click_show_list_Location.click()
        print("Successfully click show list location")
        return True, None

    except TimeoutException:
        message = "Error when click show list location"
        print(message)
        return False, message


def click_Location(self):
    try:
        # Load the "Location" value from database.json
        with open(document_path, "r") as file:
            data = json.load(file)
            location = data.get("Location", "")

        # Check if the location is valid in the JSON file
        if not location:
            print("Location value not found in database.json")
            return False, "Location value not found in database.json"

        # Determine the XPath based on the location value
        if location == "No Program":
            swipe_up_location(self)
            xpath = ('//android.widget.TextView[@text="State with No Vehicle Emission Inspection and Maintenance (I/M) '
                     'Program"]')
        elif location == "Massachusetts":
            swipe_up_location(self)
            xpath = '//android.widget.TextView[@text="Massachusetts"]'
        else:
            xpath = f'//android.widget.TextView[@text="{location}"]'

        # Click the location button based on the XPath
        location_button = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        location_button.click()
        print(f"Successfully clicked location: {location}")
        return True, None

    except TimeoutException:
        message = f"Error when clicking location: {location}"
        print(message)
        return False, message

    except Exception as e:
        message = f"Unexpected error: {e}"
        print(message)
        return False, message
