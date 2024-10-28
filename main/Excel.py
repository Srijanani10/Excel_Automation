import openpyxl
from openpyxl.drawing.image import Image
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from PIL import Image as PILImage  # Import Pillow Image
import os
import time

# Function to take a screenshot of a website after logging in
def take_screenshot(url, output_path):
    options = webdriver.EdgeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--user-data-dir=C:/temp/edge_profile")  # Use a temporary profile
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)

    driver.get(url)

    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))  # Wait for the body tag to ensure the page is loaded
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
        cropped_img = img.crop(crop_box)  # Crop the image
        cropped_img.save(output_path)  # Save the cropped image
        print(f"Cropped image saved to {output_path}")

# Function to create a copy of a sheet and add screenshots
def add_screenshots_to_excel(file_path, sheet_name, new_sheet_name, screenshots, cell_positions):
    wb = openpyxl.load_workbook(file_path)
    original_sheet = wb[sheet_name]
    new_sheet = wb.copy_worksheet(original_sheet)
    new_sheet.title = new_sheet_name

    for screenshot, cell_position in zip(screenshots, cell_positions):
        if os.path.exists(screenshot):
            img = Image(screenshot)  # Load the screenshot
            img.anchor = cell_position  # Set the cell position for the image

            column_letter = cell_position[0]
            cell_width = new_sheet.column_dimensions[column_letter].width
            cell_height = new_sheet.row_dimensions[int(cell_position[1:])].height

            if cell_width is None:
                cell_width = 8.43
            if cell_height is None:
                cell_height = 15.00

            cell_width_pixels = cell_width * 7
            cell_height_pixels = cell_height * 0.75

            img.width = cell_width_pixels
            img.height = cell_height_pixels

            new_sheet.add_image(img)

        else:
            print(f"Screenshot file not found: {screenshot}")

    wb.save(file_path)

# Define your parameters
file_path = r"C:\Users\srijanani.LTPL\Downloads\LX 70 Project Health Card Template.xltx"  # Path to the Excel file
sheet_name = 'Dashboard'  # Update this to the actual name of the sheet you want to copy
new_sheet_name = 'Dashboard Copy'  # Name for the new sheet

# URLs to capture after login
urls = [
    'https://projects.zoho.in/portal/lectrixtech#taskreports/171358000000548009/basicreports/status/customview/171358000000082007/donut',  # Replace with actual URLs
    'https://projects.zoho.in/portal/lectrixtech#taskreports/171358000000548009/basicreports/owner/customview/171358000000082007/bar',
    'https://projects.zoho.in/portal/lectrixtech#bugreports/171358000000548009/basicreports/status/customview/171358000001439105/donut',
    'https://projects.zoho.in/portal/lectrixtech#bugreports/171358000000548009/advancedreports/dynamicreport/owners/status/customview/171358000001439095/stacked'
]

# Create a list to hold the paths of screenshots
screenshot_paths = []
for i, url in enumerate(urls):
    screenshot_path = f"C:\\Git_Projects\\Excel_Automation\\screenshot_{i + 1}.png"  # Define a unique filename for each screenshot
    take_screenshot(url, screenshot_path)  # Take screenshot

    # Define cropping box (left, upper, right, lower)
    # Define cropping box (left, upper, right, lower)
    crop_box = (380, 237, 1863, 900)  # Adjust these values based on your image dimensions and desired crop area
    cropped_screenshot_path = f"C:\\Git_Projects\\Excel_Automation\\cropped_screenshot_{i + 1}.png"
    crop_image(screenshot_path, cropped_screenshot_path, crop_box)  # Crop the image
    screenshot_paths.append(cropped_screenshot_path)  # Add cropped path to the list

cell_positions = ['B4', 'C4', 'B6', 'C6']  # Corresponding cell positions for the screenshots

# Call the function
add_screenshots_to_excel(file_path, sheet_name, new_sheet_name, screenshot_paths, cell_positions)
