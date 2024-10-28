import openpyxl
from openpyxl.drawing.image import Image
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import os
import time  # Import time module for sleep

# Function to take a screenshot of a website after logging in
def take_screenshot(url, output_path):
    # Set up the Selenium WebDriver (Edge)
    options = webdriver.EdgeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--user-data-dir=C:/temp/edge_profile")  # Use a temporary profile
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)

    # Open the website
    driver.get(url)

    try:
        # Wait for a specific element on the page that indicates login was successful
        # Adjust the locator as needed to match a relevant element
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))  # Wait for the body tag to ensure the page is loaded
        )

        # Optional: Wait an additional few seconds for full page load (can be adjusted)
        time.sleep(5)

        # Take a screenshot
        driver.save_screenshot(output_path)
        print(f"Screenshot saved to {output_path}")

    except Exception as e:
        print(f"Error occurred: {e}")

    finally:
        # Close the browser
        driver.quit()

# Function to create a copy of a sheet and add screenshots
def add_screenshots_to_excel(file_path, sheet_name, new_sheet_name, screenshots, cell_positions):
    # Load the existing workbook
    wb = openpyxl.load_workbook(file_path)

    # Get the sheet to copy
    original_sheet = wb[sheet_name]

    # Create a copy of the sheet
    new_sheet = wb.copy_worksheet(original_sheet)
    new_sheet.title = new_sheet_name

    # Insert screenshots into specified cells
    for screenshot, cell_position in zip(screenshots, cell_positions):
        # Check if the screenshot file exists before trying to add it
        if os.path.exists(screenshot):
            img = Image(screenshot)  # Load the screenshot
            img.anchor = cell_position  # Set the cell position for the image

            # Get the cell width and height in pixels
            column_letter = cell_position[0]  # Get the column letter (e.g., 'A', 'B', etc.)
            cell_width = new_sheet.column_dimensions[column_letter].width  # Get column width in characters
            cell_height = new_sheet.row_dimensions[int(cell_position[1:])].height  # Get row height in points

            # Convert width and height to pixels
            if cell_width is None:
                cell_width = 8.43  # Default width if None (approx. 64 pixels)
            if cell_height is None:
                cell_height = 15.00  # Default height if None (approx. 20 pixels)

            # Convert column width (character width) to pixels
            cell_width_pixels = cell_width * 7  # Approx. 7 pixels per character
            cell_height_pixels = cell_height * 0.75  # Approx. 0.75 pixels per point

            # Set the image size to fit the cell size
            img.width = cell_width_pixels
            img.height = cell_height_pixels

            new_sheet.add_image(img)  # Add the image to the new sheet
        else:
            print(f"Screenshot file not found: {screenshot}")

    # Save the workbook with the new sheet and images
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
    screenshot_paths.append(screenshot_path)  # Add path to the list

cell_positions = ['B4', 'C4', 'B6', 'C6']  # Corresponding cell positions for the screenshots

# Call the function
add_screenshots_to_excel(file_path, sheet_name, new_sheet_name, screenshot_paths, cell_positions)
