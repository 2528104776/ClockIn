!#/usr/bin/python
import cv2,requests,sys,os,time,smtplib
from PIL import ImageFont, ImageDraw, Image
import numpy as np
from openpyxl import Workbook,load_workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

wr = load_workbook('任务清单.xlsx')
sheet = wr.active
to_name = sheet['E2'].value #邮箱
total_num = 0



def get_mission():
    global total_num
    mission = []
    for index,row in enumerate(sheet):
        if index !=0 and row[0].value is not None:
            # print(index,row[0].value)
            mission.append(row[0].value)
    if len(mission) ==0:
        print('请先在表格内添加任务!')
        input('按任意键退出')
        sys.exit(0)


    elif len(mission) ==1:
        print(f'您添加了一个任务:{mission[0]}')

        quantity = int(input("请输入本次打卡数量,(可不选,按回车跳过):"))
        if quantity:
            sheet['D2'].value += quantity
            total_num = sheet['D2'].value

        return mission[0]

    elif len(mission) > 1:
        for index,i in enumerate(mission):
            print(index+1,i)
        msg = int(input("您有多个任务,请输入数字选择:"))
        quantity = int(input("请输入本次打卡数量,(可不选,按回车跳过):"))
        if quantity:
            sheet['D2'].value += quantity
        total_num = sheet['D2'].value
        return mission[msg - 1]





def get_day(thing):
    day = 0
    now_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
    for index, row in enumerate(sheet):
        if row[0].value == thing:
            day += row[1].value
            row[1].value+=1
            row[2].value = now_time
    print(f'上次打卡天数为:{day}')
    print(f'本次打卡第{day+1}天,{thing}的打卡数量累计:{total_num}个.')

    wr.save('任务清单.xlsx')

    return day+1

def parse():
    url = 'https://rest.shanbay.com/api/v2/quote/quotes/today/'
    res = requests.get(url)
    # print(res.json())
    img = res.json()['data']['origin_img_urls'][0]
    content = res.json()['data']['content']
    translation = res.json()['data']['translation']
    image = requests.get(img).content
    with open('wallpaper.jpg','wb')as file:
        file.write(image)
    return content,translation
def image(num,thing = '背单词'):
    s1,s2 = parse()
    # 编辑图片路径
    bk_img = cv2.imread("wallpaper.jpg")
    # 设置需要显示的字体
    fontpath = "simhei.ttf"
    # 32为字体大小
    font = ImageFont.truetype(fontpath, 50)
    img_pil = Image.fromarray(bk_img)
    draw = ImageDraw.Draw(img_pil)
    # 绘制文字信息
    # (100,300/350)为字体的位置，(255,255,255)为白色，(0,0,0)为黑色
    # -------------------字符处理--------------------------

    str2 = list(s1)
    str2.insert(42,"\n")
    text1= "".join(str2)


    str5 = list(s2)
    str5.insert(21,"\n")
    text2= "".join(str5)



    draw.text((0, 0), thing, font=ImageFont.truetype(fontpath, 28), fill=(255, 255, 255),troke_width = 1)
    #pil中的颜色是BGR而不是RGB
    draw.text((350, 200), "坚持打卡天数", font=ImageFont.truetype(fontpath, 70), fill=(193,182,255),stroke_width = 1)
    draw.text((460, 290), "%2d天"%(num), font=ImageFont.truetype(fontpath, 70), fill=(193,182,255),stroke_width = 1)

    draw.text((0, 1700), text1, font=font, fill=(255, 255, 255))
    # y轴间隔100坐标,50个坐标换一行
    draw.text((0, 1800), text2, font=font, fill=(255, 255, 255))

    draw.text((800, 1630), f"加好友一起{thing}", font=ImageFont.truetype(fontpath, 28), fill=(255,255,255))
    draw.text((0, 1600), f"{thing}累计打卡数量:{total_num}个", font=ImageFont.truetype(fontpath, 50), fill=(193,255,193))



    img1 = Image.open("./imgserver.jpg")
    img_pil.paste(img1,(800,1350))


    bk_img = np.array(img_pil)


    # 保存图片路径
    cv2.imwrite("new_img.jpg", bk_img)
    os.system("del wallpaper.jpg")
    os.system("new_img.jpg")

    # os.system(r'del C:\Users\Administrator\Desktop\new_img.jpg')
    return num




def send(to_name):
    _to = to_name
    _user = "2528104776@qq.com"
    _pwd = "nxpkdslwppacdjcc"
    _to = "2528104776@qq.com"

    # 如名字所示Multipart就是分多个部分
    msg = MIMEMultipart()
    msg["Subject"] = "don't panic"
    msg["From"] = _user
    msg["To"] = _to

    # ---这是文字部分---
    part = MIMEText("恭喜打卡成功！")
    msg.attach(part)

    # jpg类型附件
    part = MIMEApplication(open('new_img.jpg', 'rb').read())
    part.add_header('Content-Disposition', 'attachment', filename="每日打卡.jpg")
    msg.attach(part)

    smtp = smtplib.SMTP_SSL("smtp.qq.com", 465)

    # HELO 向服务器标识用户身份
    smtp.helo("smtp.qq.com")
    # 服务器返回结果确认
    smtp.ehlo("smtp.qq.com")
    # 登录邮箱服务器用户名和密码
    smtp.login(_user, _pwd)
    print("Start send email...")
    smtp.sendmail(_user, _to, msg.as_string())
    smtp.quit()
    print("发送成功！")

def main():
    thing = get_mission()
    num = get_day(thing)
    parse()
    image(num,thing)
    send()

if __name__=="__main__":
    main()
    input("按任意键退出")





