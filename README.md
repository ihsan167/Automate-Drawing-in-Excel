# \# Automate Drawing in Excel

# 

# Convert any image into a pixel-art drawing made entirely from colored Excel cells,

# using Python to extract pixel RGB values and a VBA macro to color the cells.

# 

# \## How it works

# 

# 1\. \*\*Python script\*\* (`pixelValue2.py`) opens an image you choose, reads every

# &#x20;  pixel's RGB values, and writes them into the `pixelResult` sheet of

# &#x20;  `AutomateDraw.xlsm`.

# 2\. \*\*VBA macro\*\* inside `AutomateDraw.xlsm` reads those RGB values and sets each

# &#x20;  cell's background color to recreate the image visually.

# 

# See `Sample Results.xlsx` for an example of the output.

# 

# \## Prerequisites

# 

# \- \*\*Python 3.8+\*\*

# \- \*\*Windows\*\* with \*\*Microsoft Excel\*\* installed (xlwings requires Excel)

# \- Excel macros must be enabled when you open `AutomateDraw.xlsm`

# 

# \## Installation

# 

# pip install -r requirements.txt

# 

# \## Usage

# 

# 1\. Place `pixelValue2.py` and `AutomateDraw.xlsm` in the same folder.

# 2\. Run the script: `python pixelValue2.py`

# 3\. A file dialog will open — select a JPG or PNG image.

# 4\. The script writes the pixel data into `AutomateDraw.xlsm` and saves it.

# 5\. Open `AutomateDraw.xlsm` in Excel, enable macros, and run the coloring macro

# &#x20;  to render the image.

# 

# \## Notes

# 

# \- \*\*Image size:\*\* Excel supports up to 1,048,576 rows, so images must have

# &#x20; fewer than 1,048,576 pixels in total (e.g., a 1024x1024 image is fine).

# &#x20; Use a resized/thumbnail version for large photos.

# \- \*\*Supported formats:\*\* JPG, JPEG, PNG.

# \- \*\*Alpha channel:\*\* Transparent images are converted to RGB automatically

# &#x20; (transparency is discarded).



