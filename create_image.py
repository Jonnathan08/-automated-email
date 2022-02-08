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
    font_text = ImageFont.truetype('arial.ttf', 64)
    font_text2 = ImageFont.truetype('arial.ttf', 80)

    #Draw Name, date and errors number
    draw.text((65, 50), pres % name ,(68, 84, 106), font=font_text)
    draw.text((60, 320), date ,(68, 84, 106), font=font_text)
    draw.text((1000, 320), str(number) ,(68, 84, 106), font=font_text2)

    resized_img = im.resize((711,400))
    resized_img.save('imagedraw.png', width='400px')

    #im.show()

generate_image('Name', 'date', 1)