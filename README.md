
# Excel Image Extractor

## Overview

The Excel Image Extractor is a Python script that extracts all images from an Excel file and saves them in order of appearance. Each image is saved with a filename indicating its order and position in the Excel file.

## Requirements

- Python 3.x
- `openpyxl` library
- `Pillow` library

## Installation

1. Install Python 3 from [Python's official website](https://www.python.org/).
2. Install the required Python libraries using pip:

   ```bash
   pip install openpyxl Pillow
   ```

## Usage

1. Save the script (`extract_images.py`) and place it in a directory.
2. Modify the script to specify the path to your Excel file and the output directory:

   ```python
   # Define the path to your Excel file and the output folder
   excel_file_path = "path_to_your_excel_file.xlsx"
   output_directory = "output_images"
   ```

3. Run the script:

   ```bash
   python extract_images.py
   ```

## Script

```python
import openpyxl
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import os

def extract_images(excel_path, output_folder):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(excel_path, data_only=True)

    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    image_count = 1

    # Iterate over all sheets in the workbook
    for sheet in workbook.sheetnames:
        worksheet = workbook[sheet]
        
        # Iterate over all images in the sheet
        for image in worksheet._images:
            image_ref = image.anchor._from
            image_position = f"{image_ref.col}{image_ref.row}"

            # Extract the image
            pil_image = PILImage.open(image.ref)
            
            # Save the image with a name indicating its position
            image_filename = os.path.join(output_folder, f"image_{image_count}_{image_position}.png")
            pil_image.save(image_filename)

            image_count += 1

# Define the path to your Excel file and the output folder
excel_file_path = "path_to_your_excel_file.xlsx"
output_directory = "output_images"

# Extract images
extract_images(excel_file_path, output_directory)
```
