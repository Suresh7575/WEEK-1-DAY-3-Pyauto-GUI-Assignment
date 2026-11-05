
import os
import re
import sys
import time
import pyautogui
import pyperclip

from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ==================================================
# 1. CONFIGURATION
# ==================================================

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

PROJECT_FOLDER = os.path.dirname(os.path.abspath(__file__))

TODAY = datetime.now()

DATE_TEXT = TODAY.strftime("%Y-%m-%d")
TIME_TEXT = TODAY.strftime("%Y-%m-%d %H:%M:%S")

EXCEL_FILE = os.path.join(
    PROJECT_FOLDER,
    f"daily_report_{DATE_TEXT}.xlsx"
)

SCREENSHOT_FILE = os.path.join(
    PROJECT_FOLDER,
    f"daily_report_screenshot_{DATE_TEXT}.png"
)

WEBSITE_URL = (
    "https://www.timeanddate.com/weather/"
    "singapore/singapore"
)


# ==================================================
# 2. CHECK SCREENSHOT DEPENDENCIES
# ==================================================

def check_screenshot_dependencies():

    print("\nChecking screenshot dependencies...")

    try:

        from PIL import Image
        import pyscreeze

        print("Pillow:", Image.__version__ if hasattr(Image, "__version__") else "Ready")
        print("PyScreeze: Ready")
        print("PyAutoGUI: Ready")

        return True

    except ImportError as error:

        print("Screenshot dependency missing:", error)

        print("\nPlease run this command in CMD:")
        print(
            "python -m pip install --upgrade "
            "Pillow pyscreeze pyautogui"
        )

        return False


# ==================================================
# 3. OPEN CHROME
# ==================================================

def open_chrome():

    print("\n[1/6] Opening Google Chrome...")

    try:

        pyautogui.hotkey("win", "r")
        time.sleep(2)

        pyautogui.write(
            f'chrome "{WEBSITE_URL}"',
            interval=0.04
        )

        pyautogui.press("enter")

        print("Chrome launched successfully.")

        # Wait for website loading
        time.sleep(15)

        return True

    except Exception as error:

        print("Chrome opening error:", error)

        return False


# ==================================================
# 4. COPY WEBSITE INFORMATION
# ==================================================

def copy_website_data():

    print("\n[2/6] Collecting website information...")

    try:

        screen_width, screen_height = pyautogui.size()

        # Focus browser window
        pyautogui.click(
            screen_width // 2,
            screen_height // 2
        )

        time.sleep(2)

        # Select webpage content
        pyautogui.hotkey("ctrl", "a")

        time.sleep(1)

        # Copy selected content
        pyautogui.hotkey("ctrl", "c")

        time.sleep(3)

        # Read clipboard
        page_content = pyperclip.paste()

        if not page_content:

            raise ValueError(
                "No webpage content was copied."
            )

        print("Website information copied successfully.")

        return page_content

    except Exception as error:

        print("Data collection error:", error)

        return None


# ==================================================
# 5. EXTRACT WEATHER INFORMATION
# ==================================================

def extract_information(page_content):

    print("\n[3/6] Extracting weather information...")

    temperature_match = re.search(
        r"(-?\d{1,2})\s*°\s*C",
        page_content,
        re.IGNORECASE
    )

    if temperature_match:

        temperature = temperature_match.group(0)

        fetched_data = (
            f"Singapore Current Temperature: {temperature}"
        )

        comment = (
            "Weather information collected successfully."
        )

    else:

        fetched_data = (
            "Temperature not found on website."
        )

        comment = (
            "Please verify the website manually."
        )

    print("Fetched information:", fetched_data)

    return fetched_data, comment


# ==================================================
# 6. CREATE EXCEL REPORT
# ==================================================

def create_excel_report(fetched_data, comment):

    print("\n[4/6] Creating Excel report...")

    try:

        workbook = Workbook()

        sheet = workbook.active

        sheet.title = "Daily Operations Report"

        # Header columns
        headers = [
            "Date & Time",
            "Fetched Information",
            "Operations Comment"
        ]

        sheet.append(headers)

        # Report information
        sheet.append([
            TIME_TEXT,
            fetched_data,
            comment
        ])

        # Style headers
        for cell in sheet[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF",
                size=11
            )

            cell.fill = PatternFill(
                start_color="1F4E78",
                end_color="1F4E78",
                fill_type="solid"
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # Style report data
        for cell in sheet[2]:

            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

        # Column widths
        column_widths = [25, 55, 45]

        for index, width in enumerate(
            column_widths, start=1
        ):

            sheet.column_dimensions[
                get_column_letter(index)
            ].width = width

        sheet.row_dimensions[1].height = 25
        sheet.row_dimensions[2].height = 40

        # Freeze header row
        sheet.freeze_panes = "A2"

        # Save Excel file
        workbook.save(EXCEL_FILE)

        if not os.path.exists(EXCEL_FILE):

            raise FileNotFoundError(
                "Excel file was not created."
            )

        print("Excel report created successfully!")

        print("Excel location:", EXCEL_FILE)

        return True

    except Exception as error:

        print("Excel creation error:", error)

        return False


# ==================================================
# 7. OPEN EXCEL REPORT
# ==================================================

def open_excel_report():

    print("\n[5/6] Opening Excel report...")

    try:

        os.startfile(EXCEL_FILE)

        print("Excel report opened successfully.")

        # Wait for Excel to open completely
        time.sleep(10)

        return True

    except Exception as error:

        print("Excel opening error:", error)

        return False


# ==================================================
# 8. TAKE SCREENSHOT
# ==================================================

def take_screenshot():

    print("\n[6/6] Taking final screenshot...")

    try:

        # Import screenshot module
        from PIL import Image
        import pyscreeze

        # Capture complete desktop
        screenshot = pyautogui.screenshot()

        # Save screenshot
        screenshot.save(SCREENSHOT_FILE, format="PNG")

        # Verify file
        if not os.path.isfile(SCREENSHOT_FILE):

            raise FileNotFoundError(
                "Screenshot file was not generated."
            )

        # Check file size
        file_size = os.path.getsize(SCREENSHOT_FILE)

        if file_size == 0:

            raise ValueError(
                "Screenshot file is empty."
            )

        # Verify image can be opened
        with Image.open(SCREENSHOT_FILE) as image:

            image.verify()

        print("\nScreenshot saved successfully!")

        print("Screenshot location:")
        print(SCREENSHOT_FILE)

        print(
            f"Screenshot size: {file_size:,} bytes"
        )

        return True

    except Exception as error:

        print("\nSCREENSHOT ERROR:", error)

        print("\nPlease check these libraries:")
        print("Pillow")
        print("PyScreeze")
        print("PyAutoGUI")

        return False


# ==================================================
# 9. FINAL STATUS
# ==================================================

def show_final_status(screenshot_success):

    print("\n" + "=" * 50)

    print("       DAILY REPORT BOT COMPLETED")

    print("=" * 50)

    print("\nExcel report:")
    print(EXCEL_FILE)

    print("\nScreenshot:")
    print(SCREENSHOT_FILE)

    excel_status = os.path.exists(EXCEL_FILE)

    print("\nExcel created:", excel_status)

    print("Screenshot created:", screenshot_success)

    if excel_status and screenshot_success:

        print("\nSUCCESS: ALL FILES GENERATED!")

    else:

        print("\nWARNING: SOME FILES ARE MISSING.")

    print("\nThank you for using the automation bot!")


# ==================================================
# 10. MAIN PROGRAM
# ==================================================

def main():

    print("\n" + "=" * 50)

    print("       DAILY OPERATIONS REPORT BOT")

    print("=" * 50)

    print("Date:", DATE_TEXT)

    print("Time:", TIME_TEXT)

    # Check screenshot dependencies before automation
    if not check_screenshot_dependencies():

        return

    # Step 1: Open Chrome
    if not open_chrome():

        return

    # Step 2: Copy website information
    page_content = copy_website_data()

    if not page_content:

        return

    # Step 3: Extract information
    fetched_data, comment = extract_information(
        page_content
    )

    # Step 4: Create Excel report
    excel_success = create_excel_report(
        fetched_data,
        comment
    )

    if not excel_success:

        return

    # Step 5: Open Excel
    if not open_excel_report():

        return

    # Step 6: Take screenshot
    screenshot_success = take_screenshot()

    # Final report
    show_final_status(screenshot_success)


# ==================================================
# 11. RUN PROGRAM
# ==================================================

if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        print("\nBot stopped by user.")

    except Exception as error:

        print("\nUnexpected error:", error)

        sys.exit(1)
