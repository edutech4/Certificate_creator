import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from PyPDF2 import PdfFileMerger, PdfFileReader
import os
import shutil

import qrcode


#--------------------------------------------------------------------------------------------------------------------
#-------------THIS SECTION OF CODE GENERATES QR-CODE AND SAVES AS AN IMAGE WCHICH WILL BE ATTACHED TO THE CERTIFICATE--
# image = qrcode.make(input("Write Your message: "))
# image.save("qrcode.jpg")


#--------------------------------------------------------------------------------------------------------------------
#-------------THIS SECTION OF CODE ATTACHES NAMES GOTTEN FROM EXCEL TO THE TEMPLATE OF THE CERTIFICATE---------------
data = pd.read_excel(r'C:/Users/MO_WIZZ/Documents/Python_Project/Certificate_creator/certname_giz.xlsx')
name_list = data["CERTNAME"].tolist()
# name_date = data["DATE"].tolist()
# qr_code_data = data["CERTNO"].tolist()
# print(name_list)
# print(name_date[0])
# print(qr_code_data)
# image = qrcode.make(qr_code_data[0])
# image.save("qrcode.jpg")
file_number = 0
number_of_files = 0
for i in name_list:
    i_edited = i.replace(" ","  ")  #Adding more spaces between surname and firstname
    # print(i)
    # i_edited = name_list[5]
    # print(len(name_list[5]))
    im = Image.open(r'C:/Users/MO_WIZZ/Documents/Python_Project/Certificate_creator/GIZ_CERT_template.jpg')
    d = ImageDraw.Draw(im)
    # location = (1150, 1070)  #x, y Robotics centre certificate cordinate
    if len(i_edited) < 18:
        location = (1350, 1254)  # x, y
    elif len(i_edited) > 21:
        location = (1000, 1254)  # x, y
    else:
        location = (1230, 1254)  # x, y
    text_color = (0, 137, 209)
    font = ImageFont.truetype("ITCBLKAD.ttf", 150)
    d.text(location, i_edited, fill =29, font=font)
    file_number = file_number + 1
    number_of_files = file_number
    # im.save("certificate_" + i + ".pdf")  #Saves as pdf
    im.save(str(file_number) + ".pdf")  # Saves as pdf

##------------------THIS SECTION HELPS TO MERGE MULTIPLE PDF FILES-------------------------------------------------------
file_number = 0
mergedObject = PdfFileMerger()
for fileNumber in range(number_of_files):
    file_number = file_number + 1
    mergedObject.append(PdfFileReader(str(file_number) + '.pdf', 'rb'))

# Write all the files into a file which is named as shown below
mergedObject.write("mergedfilesoutput.pdf")


#####--------------------THIS SECTION HELPS TO DELETE MULTIPLE FILES THAT WAS CREATED BEFORE THIS SECTION-------------------
# for filename in range(1, 51):
#     if os.path.exists(str(filename)+'.pdf'):
#         os.remove(str(filename)+'.pdf')

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



