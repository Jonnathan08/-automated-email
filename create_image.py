from turtle import width
from PIL import Image, ImageDraw, ImageFont
from numpy import size

def generate_image(name, date, number):    
    
    pres = '    %s'
    #Import template
    im = Image.open('template.jpg')

    #create label with the import image
    draw = ImageDraw.Draw(im)

    #set font
    font_text = ImageFont.truetype('pixel.ttf', 80)
    font_text2 = ImageFont.truetype('pixel.ttf', 80)

    #Draw Name, date and errors number
    draw.text((65, 105), pres % name ,(68, 84, 106), font=font_text)
    draw.text((230, 630), date ,(237, 59, 59), font=font_text)
    draw.text((1100, 630), str(number) ,(237, 59, 59), font=font_text)

    resized_img = im.resize((700,393))
    resized_img.save('imagedraw.png', width='400px')

    #im.show()

def generate_image_congrat():    
    
    #Import template
    im = Image.open('template2.jpg')

    resized_img = im.resize((700,496))
    resized_img.save('imagedraw.png', width='400px')


generate_image_congrat()