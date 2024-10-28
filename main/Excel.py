import openpyxl
from openpyxl.drawing.image import Image
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from PIL import Image as PILImage
import os
import time

# Function to take a screenshot of a website after logging in
def take_screenshot(url, output_path):
    options = webdriver.EdgeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--user-data-dir=C:/temp/edge_profile")
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)

    driver.get(url)

    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        time.sleep(5)
        driver.save_screenshot(output_path)
        print(f"Screenshot saved to {output_path}")

    except Exception as e:
        print(f"Error occurred: {e}")

    finally:
        driver.quit()

# Function to crop the image
def crop_image(input_path, output_path, crop_box):
    with PILImage.open(input_path) as img:
        cropped_img = img.crop(crop_box)
        cropped_img.save(output_path)
        print(f"Cropped image saved to {output_path}")

# Function to create a copy of a sheet and add screenshots with custom dimensions
def add_screenshots_to_excel(file_path, sheet_name, new_sheet_name, screenshots, cell_positions):
    wb = openpyxl.load_workbook(file_path)
    original_sheet = wb[sheet_name]
    new_sheet = wb.copy_worksheet(original_sheet)
    new_sheet.title = new_sheet_name

    # Define custom column widths and row heights for specific cells
    dimensions = {
        'B4': {'width': 74.75, 'height': 148},
        'C4': {'width': 96.75, 'height': 148},
        'B6': {'width': 74.75, 'height': 168},
        'C6': {'width': 96.75, 'height': 168}
    }

    for screenshot, cell_position in zip(screenshots, cell_positions):
        if os.path.exists(screenshot):
            # Adjust column width and row height
            col_letter = cell_position[0]
            row_number = int(cell_position[1:])

            if cell_position in dimensions:
                new_sheet.column_dimensions[col_letter].width = dimensions[cell_position]['width']
                new_sheet.row_dimensions[row_number].height = dimensions[cell_position]['height']

            # Load and resize the screenshot
            img = Image(screenshot)
            img.width = new_sheet.column_dimensions[col_letter].width * 7.5
            img.height = new_sheet.row_dimensions[row_number].height * 0.75
            img.anchor = cell_position

            new_sheet.add_image(img)
        else:
            print(f"Screenshot file not found: {screenshot}")

    wb.save(file_path)

# Define parameters
file_path = r"C:\Users\srijanani.LTPL\Downloads\LX 70 Project Health Card Template.xltx"
sheet_name = 'Dashboard'
new_sheet_name = 'Dashboard Copy'

urls = [
    'https://projects.zoho.in/portal/lectrixtech#taskreports/171358000000548009/basicreports/status/customview/171358000000082007/donut',
    'https://projects.zoho.in/portal/lectrixtech#taskreports/171358000000548009/basicreports/owner/customview/171358000000082007/bar',
    'https://projects.zoho.in/portal/lectrixtech#bugreports/171358000000548009/basicreports/status/customview/171358000001439105/donut',
    'https://projects.zoho.in/portal/lectrixtech#bugreports/171358000000548009/advancedreports/dynamicreport/owners/status/customview/171358000001439095/stacked'
]

screenshot_paths = []
for i, url in enumerate(urls):
    screenshot_path = f"C:\\Git_Projects\\Excel_Automation\\screenshot_{i + 1}.png"
    take_screenshot(url, screenshot_path)

    crop_box = (380, 237, 1863, 900)
    cropped_screenshot_path = f"C:\\Git_Projects\\Excel_Automation\\cropped_screenshot_{i + 1}.png"
    crop_image(screenshot_path, cropped_screenshot_path, crop_box)
    screenshot_paths.append(cropped_screenshot_path)

cell_positions = ['B4', 'C4', 'B6', 'C6']

# Call the function
add_screenshots_to_excel(file_path, sheet_name, new_sheet_name, screenshot_paths, cell_positions)
