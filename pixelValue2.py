from PIL import Image

from pandas import DataFrame

import os 

from tkinter import filedialog

im = Image.open(filedialog.askopenfilename(title = 'Ihsan Kacak sangat ni', filetypes = [('JPG','*.jpg'),('JPG','*.png'),('JPG','*.PNG')]))
#im = Image.open(r'D:\Application\mencuba.jpg')
im.show()
rgbIm = im.convert('RGB')
#im.show()



imageX,imageY = im.size
print(imageX)
print(imageY)

pix_val = list(rgbIm.getdata()) #this is tuple in list

#pix_val is a list


pix_val_flat = [x for sets in pix_val for x in sets] #This is integer in list



Rvalue = [R[0] for R in pix_val]
Gvalue = [G[1] for G in pix_val]
Bvalue = [B[2] for B in pix_val]
#print(len(Rvalue))
print((Gvalue[26]))
#print(len(Bvalue))

#print(len(pix_val))

print(pix_val[26]) 

#print(Rvalue[168520])
#kenapa ihsan tu kacak sangat

#This part is to write the pixel value into csv
pixValue = {'Red Value': Rvalue,'Green Value': Rvalue,'Blue Value' :Bvalue,'XWidth':imageX, 'YHeight':imageY}
RGBColor = list(pixValue.keys())

print  (RGBColor)

current = os.getcwd()

df = DataFrame(pixValue, columns = [RGBColor[0],RGBColor[1], RGBColor[2],RGBColor[3],RGBColor[4]])

export_csv = df.to_csv(current +  r'\pixelResult1.csv', index = True, header =True)

