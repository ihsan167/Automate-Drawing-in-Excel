from PIL import Image
import pandas as pd
import xlwings as xw
from tkinter import filedialog

EXCEL_MAX_PIXELS = 1_048_576  # Excel row limit


def main():
    # Prompt user to select an image file
    image_path = filedialog.askopenfilename(
        title='Select an image file',
        filetypes=[('Image Files', '*.jpg *.jpeg *.png *.PNG *.JPEG')]
    )
    if not image_path:
        print("No file selected. Exiting.")
        return

    # Open and convert image
    print("Loading image...")
    try:
        im = Image.open(image_path)
    except Exception as e:
        print(f"Error opening image: {e}")
        return

    rgbIm = im.convert('RGB')
    imageX, imageY = im.size
    print(f"Image size: {imageX}x{imageY} pixels ({imageX * imageY} total pixels)")

    if imageX * imageY > EXCEL_MAX_PIXELS:
        print(
            f"Error: Image has {imageX * imageY} pixels, which exceeds Excel's "
            f"limit of {EXCEL_MAX_PIXELS} rows. Please use a smaller image."
        )
        return

    # Extract RGB channel values from every pixel
    pix_val = list(rgbIm.getdata())
    Rvalue = [p[0] for p in pix_val]
    Gvalue = [p[1] for p in pix_val]
    Bvalue = [p[2] for p in pix_val]

    df = pd.DataFrame({
        'Red Value':   Rvalue,
        'Green Value': Gvalue,
        'Blue Value':  Bvalue,
        'XWidth':      imageX,
        'YHeight':     imageY,
    })

    # Write pixel data into the Excel workbook
    print("Opening AutomateDraw.xlsm...")
    try:
        app = xw.App(visible=False)
        wb = xw.Book('AutomateDraw.xlsm')
    except Exception as e:
        print(f"Error opening AutomateDraw.xlsm: {e}")
        print("Make sure AutomateDraw.xlsm is in the same directory as this script.")
        return

    try:
        ws = wb.sheets['pixelResult']
    except Exception as e:
        print(f"Error accessing sheet 'pixelResult': {e}")
        wb.close()
        app.quit()
        return

    print("Writing pixel data to Excel...")
    ws.range('A:F').api.Delete()
    ws.range('A1').options(index=True).value = df

    wb.save()
    wb.close()
    app.quit()
    print("Done! Open AutomateDraw.xlsm and run the macro to see your drawing.")


if __name__ == '__main__':
    main()
