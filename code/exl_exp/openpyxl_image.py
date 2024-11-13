from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import units

# Create a new workbook and select the active sheet
workbook = Workbook()
sheet = workbook.active

# Define and merge the cell range
merged_range = 'B2:E5'
sheet.merge_cells(merged_range)

# Optionally, set specific row heights and column widths for accuracy
sheet.column_dimensions['B'].width = 15
sheet.column_dimensions['C'].width = 20
sheet.column_dimensions['D'].width = 15
sheet.column_dimensions['E'].width = 20

sheet.row_dimensions[2].height = 30
sheet.row_dimensions[3].height = 25
sheet.row_dimensions[4].height = 30
sheet.row_dimensions[5].height = 25

# Load the image and set its size (optional)
img = Image('assets/logo.png')  # Replace with the path to your image
img.width = 80  # Set image width in points if needed
img.height = 60  # Set image height in points if needed

# Calculate the total width of the merged cell range in pixels
start_col = merged_range.split(':')[0][0]
end_col = merged_range.split(':')[1][0]

merged_cell_width = sum(units.points_to_pixels(sheet.column_dimensions[str(col)].width * 7.2) for col in range(ord(start_col), ord(end_col) + 1))

# Calculate the total height of the merged cell range in pixels
start_row = int(merged_range.split(':')[0][1:])
end_row = int(merged_range.split(':')[1][1:])

merged_cell_height = sum(units.points_to_pixels(sheet.row_dimensions[row].height or 15)
                         for row in range(start_row, end_row + 1))

# Calculate the center position within the merged cell range
x_offset = (merged_cell_width - img.width) / 2
y_offset = (merged_cell_height - img.height) / 2

# Add the image to the worksheet at the top-left cell of the merged range with calculated offsets
sheet.add_image(img, f'{start_col}{start_row}')
# img.anchor = f'{start_col}{start_row},{int(x_offset)},{int(y_offset)}'

# Save the workbook
workbook.save('ignore/centered_image1000.xlsx')
