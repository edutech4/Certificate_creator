import pandas as pd
from PIL import Image, ImageDraw, ImageFont
# import glob
# import os


# images = glob.glob("C:/Users/MO_WIZZ/Documents/Python_Project/Certificate_creator/cerrtificate_test.png")
#
# for img in images:
#     images = Image.open(img)
#     draw = ImageDraw.Draw(images)
#     font = ImageFont.truetype("BITCBLKAD.ttf", 20)
#     text = "Whatever text"
#     draw.text((191,193),text,(250,0,0),font=font)
#     images.save(img)

data = pd.read_excel(r'C:/Users/MO_WIZZ/Documents/Python_Project/Certificate_creator/test.xlsx')

name_list = data["NAME"].tolist()
# print(name_list)
for i in name_list:
    i_edited = i.replace(" ","    ")  #Adding more spaces between surname and firstname
    # print(i)
    im = Image.open(r'C:/Users/MO_WIZZ/Documents/Python_Project/Certificate_creator/cerrtificate_test.png')
    d = ImageDraw.Draw(im)
    location = (191, 193)  #x, y
    text_color = (0, 137, 209)
    font = ImageFont.truetype("ITCBLKAD.ttf", 16)
    d.text(location, i_edited, fill =209, font=font)
    im.save("certificate_" + i + ".pdf")



