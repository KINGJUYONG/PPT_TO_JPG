import os
from PIL import Image, ImageFont, ImageDraw
import win32com.client
from pdf2image import convert_from_path
import zipfile
import time

def exit():
    time.sleep(10000)
    print("exit. . .")

# 현재 디렉토리 파일에 pptx 파일 목록
ppt_name = [i for i in os.listdir() if i.endswith(".pptx")]

for myName in ppt_name:
    ppt_name = myName
    user_name = ppt_name[:-5]
    dir_name = os.getcwd()+'\\'
    temp_name = 'temp'
    abosol_path = os.getcwd()+'\\'

    if os.path.isdir(temp_name):
        for i in os.listdir(temp_name):
            os.remove(os.path.join(temp_name, i))
    else:
        os.mkdir(temp_name)

    try:
        powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
        powerpoint.Visible = 1
        deck = powerpoint.Presentations.Open(os.path.join(os.getcwd(), ppt_name))
        deck.SaveAs(dir_name + "temp.pdf", 32)
        deck.Close()
    except:
        print("PowerPoint load error")
        print("PPT file not found")
        exit()

    pop_path = './pop/Library/bin'

    try:
        images = convert_from_path("./temp.pdf", poppler_path=pop_path)
        for i in range(len(images)):
            file234 = dir_name + 'temp\\' + user_name + "_" + str(i+1) + ".jpg"
            images[i].save(file234, "JPEG")
            
            # ppt 각 페이지에 이름 붙이기
            if i != 0 and i != len(images) - 1:
                img = Image.open(file234)
                imgSize = img.size
                draw = ImageDraw.Draw(img)
                FFFF = ImageFont.truetype("fonts\MALGUNBD.TTF", 50)
                draw.text((imgSize[0]//2,imgSize[1]-20), user_name , fill = (0,0,0), font=FFFF, stroke_width = 4, stroke_fill = (255,255,255), anchor="mb") # x=0, y=10, (0,0,0) : 검은색(RGB값)
                draw = ImageDraw.Draw(img)
                img.save(file234, "JPEG")

        os.remove('temp.pdf')
    except:
        print('Pop error')
        exit()

    try:
        zip_file = zipfile.ZipFile('./' + user_name + ".zip", "w")
        for file in os.listdir(temp_name):
            if file.endswith('.jpg'):
                zip_file.write(os.path.join(temp_name, file), compress_type=zipfile.ZIP_DEFLATED)
        zip_file.close()
    except:
        print('zip error')
        exit()
        
    if os.path.isdir(temp_name):
        for i in os.listdir(temp_name):
            os.remove(os.path.join(temp_name, i))
        os.rmdir(temp_name)
    
powerpoint.Quit()