import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from PIL import Image as PILImage
import time
import tkinter as tk
from tkinter import filedialog, simpledialog

def automate_excel(file_path, links, row_numbers, heading):
    # Load the workbook and select the first sheet
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    # Create a copy of the first sheet
    copied_sheet = wb.copy_worksheet(sheet)
    copied_sheet.title = f"{sheet.title}_copy"

    # Set up Selenium WebDriver
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    service = Service('path/to/your/chromedriver')  # Update with the correct path to your chromedriver
    driver = webdriver.Chrome(service=service, options=chrome_options)

    for link, row in zip(links, row_numbers):
        # Navigate to the website
        driver.get(link)
        time.sleep(2)  # Wait for the page to load

        # Take a screenshot
        screenshot_path = f"screenshot_{row}.png"
        driver.save_screenshot(screenshot_path)

        # Crop the image (example: crop to 800x600 starting from top-left corner)
        img = PILImage.open(screenshot_path)
        cropped_img = img.crop((0, 0, 800, 600))
        cropped_img.save(screenshot_path)

        # Insert the image into the copied sheet
        img = openpyxl.drawing.image.Image(screenshot_path)
        copied_sheet.add_image(img, f'A{row}')

    # Save the workbook
    wb.save(file_path)

    # Close the WebDriver
    driver.quit()

# Example usage
# Example usage (commented out for GUI integration)
# file_path = 'path/to/your/excel_file.xlsx'
# links = ['http://example.com', 'http://example2.com']
# row_numbers = [2, 5]
# heading = 'Example Heading'

# automate_excel(file_path, links, row_numbers, heading)
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path

def get_number_of_pictures():
    num_pictures = simpledialog.askinteger("Input", "How many pictures do you want to take?")
    return num_pictures

def get_links_and_rows(num_pictures):
    links = []
    row_numbers = []
    for _ in range(num_pictures):
        link = simpledialog.askstring("Input", "Enter website link:")
        row = simpledialog.askstring("Input", "Enter cell (e.g., A1, B2) for this link:")
        links.append(link)
        row_numbers.append(row)
    return links, row_numbers

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    file_path = select_file()
    heading = simpledialog.askstring("Input", "Enter heading:")
    num_pictures = get_number_of_pictures()
    links, row_numbers = get_links_and_rows(num_pictures)

    automate_excel(file_path, links, row_numbers, heading)

if __name__ == "__main__":
    main()