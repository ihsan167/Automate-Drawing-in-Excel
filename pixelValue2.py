from PIL import Image
import pandas as pd
from pandas import DataFrame
import os
import xlwings as xw

from tkinter import filedialog

im = Image.open(filedialog.askopenfilename(title = 'Ihsan Kacak sangat ni', filetypes = [('JPG','*.jpg'),('JPG','*.png'),('JPG','*.PNG'),('JPEG','*.JPEG')]))

rgbIm = im.convert('RGB')


imageX,imageY = im.size
print(imageX)
print(imageY)

pix_val = list(rgbIm.getdata()) #this is tuple in list


pix_val_flat = [x for sets in pix_val for x in sets] #This is integer in list


Rvalue = [R[0] for R in pix_val]
Gvalue = [G[1] for G in pix_val]
Bvalue = [B[2] for B in pix_val]


#This part is to write the pixel value into csv
pixValue = {'Red Value': Rvalue,'Green Value': Rvalue,'Blue Value' :Bvalue,'XWidth':imageX, 'YHeight':imageY}
RGBColor = list(pixValue.keys())

print  (RGBColor)

df = DataFrame(pixValue, columns = [RGBColor[0],RGBColor[1], RGBColor[2],RGBColor[3],RGBColor[4]])
#print (df.head())
#current = os.getcwd()
app = xw.App(visible = False)
wb = xw.Book('AutomateDraw.xlsm')
ws = wb.sheets['pixelResult']
xw.Range('A:F').api.Delete()

ws.range('A1').options(index = True).value = df

for sheet in wb.sheets:
    if 'Sheet1' in sheet.name:
        sheet.delete()


#Close workbook
wb.save()
wb.close()
app.quit()


#export_csv = df.to_csv(current + r'\pixelResult1.csv', index = True, header =True)

