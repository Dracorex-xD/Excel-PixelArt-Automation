from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Color
#get the image path
image_path = "./assets/pexels-eberhard-grossgasteiger-1287145.jpg"
image = Image.open(image_path)

width, height = image.size

workbook = Workbook()
sheet = workbook.active

for y in range(height):
    for x in range(width):
        # Extract the RGB values of the pixel
        rgb = image.getpixel((x, y))
        # Convert RGB values to Excel color
        red, green, blue = rgb
        hex_color = f"{red:02X}{green:02X}{blue:02X}"
        excel_color = Color(rgb=hex_color)
        # Set the color of the corresponding cell in the Excel sheet
        cell = sheet.cell(row=y + 1, column=x + 1)
        cell.fill = excel_color

#REMEMBER TO CHANGE THIS \/
output_path = "C:/Users/Draco/Downloads/Excel/Book.xlsx"
workbook.save(output_path)