# DrawXslx
Think of each cell in excel as a pixel and draw an image on the sheet.
import xlwings as xw
from PIL import Image
import numpy as np

App = xw.App(visible=False, add_book=False)
wb = App.books.add()
sheet = wb.sheets['sheet1']

img = Image.open('#picture_name’.jpg')    #Open the picture to be drawn
array = img.load()
size = img.size

def pixel(w,h,color):   #Draw a pixel
    cell= sheet.range(w,h)
    cell.column_width = 3
    cell.row_height = 20
    cell.color = color

def Draw(a,size):    #Draw pciture   (If this picture is large, it will take a long time）
    w,h = size
    for i in range(0,w):
        for j in range(0,h):
            color = a[i,j]
            pixel(j + 1,i + 1,color)
            global num
            num += 1
            percentage = w*h/100
            percent = num / percentage
            if num%5 == 0:
                print('\rThe current progress is：{0:.2f}%'.format(percent),end= '')    #Show progress

num =0
Draw(array, size)
print('\over')
wb.save('#xlsx_name')
wb.close()
App.quit()

