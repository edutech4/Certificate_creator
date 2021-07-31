import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import qrcode


#--------------------------------------------------------------------------------------------------------------------
#-------------THIS SECTION OF CODE GENERATES QR-CODE AND SAVES AS AN IMAGE WCHICH WILL BE ATTACHED TO THE CERTIFICATE--
# image = qrcode.make(input("Write Your message: "))
# image.save("qrcode.jpg")



#--------------------------------------------------------------------------------------------------------------------
#-------------THIS SECTION OF CODE ATTACHES NAMES GOTTEN FROM EXCEL TO THE TEMPLATE OF THE CERTIFICATE---------------
data = pd.read_excel(r'C:/Users/MO_WIZZ/Documents/Python_Project/Certificate_creator/test.xlsx')
name_list = data["NAME"].tolist()
# print(name_list)
for i in name_list:
    i_edited = i.replace(" ","    ")  #Adding more spaces between surname and firstname
    # print(i)
    im = Image.open(r'C:/Users/MO_WIZZ/Documents/Python_Project/Certificate_creator/CERTIFICATE2.jpg')
    d = ImageDraw.Draw(im)
    location = (1260, 1125)  #x, y
    text_color = (0, 137, 209)
    font = ImageFont.truetype("ITCBLKAD.ttf", 90)
    d.text(location, i_edited, fill =29, font=font)
    im.save("certificate_" + i + ".pdf")  #Saves as pdf


#--------------------------------------------------------------------------------------------------------------------
#-------------THIS SECTION OF CODE ATTACHES QR CODE TO THE TEMPLATE OF THE CERTIFICATE-------------------------------

# img_to_paste = Image.open('C:/Users/MO_WIZZ/Documents/Python_Project/Certificate_creator/roboticscentre.png')
# width, height = img_to_paste.size
#
# n = 2  # paste image n-times
#
# img = Image.new('RGB', (width * n, height), color=(0, 0, 0))  # set width, height of new image based on original image
# img.save('CERT.png')
#
# out = Image.open('CERTIFICATE2.png')  #background image where other image is being attached
# out.paste(img_to_paste, (2963, 335))  # the second argument here is tuple representing upper left corner
# out.save('CERTIFICATE4.png')  #Finally generated image



