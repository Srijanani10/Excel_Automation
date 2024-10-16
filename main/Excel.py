import openpyxl
from openpyxl.drawing.image import Image

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

    # Save the workbook with the new sheet and images
    wb.save(file_path)

# Define your parameters
file_path = r"C:\Users\srijanani.LTPL\Downloads\LX 70 Project Health Card Template.xltx"  # Path to the Excel file
sheet_name = 'Dashboard'  # Update this to the actual name of the sheet you want to copy
new_sheet_name = 'Dashboard Copy'  # Name for the new sheet
screenshots = [
    r"C:\Users\srijanani.LTPL\OneDrive - Lectrix Technologies Private Limited\jun-20\Pictures\Screenshots\Screenshot 2024-10-16 113311.png",
    r"C:\Users\srijanani.LTPL\OneDrive - Lectrix Technologies Private Limited\jun-20\Pictures\Screenshots\Screenshot 2024-10-16 113401.png",
    r"C:\Users\srijanani.LTPL\OneDrive - Lectrix Technologies Private Limited\jun-20\Pictures\Screenshots\Screenshot 2024-10-16 113451.png",
    r"C:\Users\srijanani.LTPL\OneDrive - Lectrix Technologies Private Limited\jun-20\Pictures\Screenshots\Screenshot 2024-10-16 113525.png"
]  # Paths to the screenshots
cell_positions = ['B4', 'C4', 'B6', 'C6']  # Corresponding cell positions for the screenshots

# Call the function
add_screenshots_to_excel(file_path, sheet_name, new_sheet_name, screenshots, cell_positions)
