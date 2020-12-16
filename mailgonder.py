"""
Author: Emir Polat
Date: 21.11.2020
"""

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import xlrd
import time, sys
from colorama import init
from colorama import Fore, Back, Style
from getpass import getpass

init(autoreset=True)

try:
    file = input(Fore.GREEN + "Excel Dosyasının Yolunu Belirtiniz: ")

    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)

    firstname = []
    emails = []
    gender = []
    # Listedeki bütün her şeyi getirerek bir listeye atamak için yapılan döngü
    for i in range(sheet.nrows):
        name = sheet.row_values(i)
        all_emails = name[3]
        all_first_name = name[0]
        all_surname = name[1]
        all_gender = name[2]

        firstname.append(all_first_name)
        gender.append(all_gender)
        emails.append(all_emails)
except Exception as e:
    print(Back.RED + Fore.WHITE +  f"Excel dosyasını ararken bir hata meydana geldi ! Lütfen geliştirici ile iletişime geçin ! \n Hata kodu: {e}")
    sys.exit()


def send_mail(gender, mesajlar, subject):
    global msg

file = ("Book.xlsx")


    try:
        msg = MIMEMultipart()
        s = smtplib.SMTP(host="SMTP.office365.com", port=587)
        s.starttls()
        s.login(req, passw)

        msg['From'] = req
        msg['To'] = emails[0]
        msg['Subject'] = mesajlar

        msg.attach(MIMEText(gender+message, 'plain'))

        if check == "E" or check == "e":
            send_image()
        else:
            pass

        print("\n\n" + Fore.CYAN +emails[0] + Fore.LIGHTGREEN_EX + " -> Mail Başarıyla Gönderildi")
        s.send_message(msg)

        del msg

    except Exception as e:
        print(Back.RED + Fore.WHITE +f"Bir hata meydana geldi ! Lütfen geliştirici ile iletişime Geçin ! \n Hata kodu: {e}")

def send_image():

    try:
        attachment = open(resim, "rb")
        p = MIMEBase('application', 'octet-stream')
        p.set_payload((attachment).read())
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment; filename= %s" % resim)
        msg.attach(p)

    except Exception as e:
        print(Back.RED + Fore.WHITE +f"Resim ekleme ya da yollamadan kaynaklı bir hata meydana geldi ! Lütfen geliştirici ile iletişime geçiniz ! \n Hata kodu: {e}")

req = input(Fore.GREEN +"Mail Adresinizi Giriniz: ")
passw = getpass(Fore.GREEN +"Parola: ")

subject = input(Fore.GREEN + "Göndereceğiniz Mail Başlığını Giriniz: ")
try:
    message = open('mailicerik.txt', 'r').read()
except Exception as e:
    print(f"Mail İçerik Dosyasını Açarken Hata Meydana Geldi ! Lütfen Geliştirici İle İletişime Geçin ! \n Hata Kodu: {e}")

check = input( Fore.GREEN + "Resim Olsun Mu ? (E/h)")

if check == "E" or check == "e":
    resim = input(Fore.GREEN + "Resim Yolunu Giriniz:")
else:
    pass

while gender:

    if gender[0] == "E":
        man_message = f"Merhaba {firstname[0]} Bey, \n\n"
        send_mail(man_message, subject, message)
    else:
        woman_message = f"Merhaba {firstname[0]} Hanım, \n\n"
        send_mail(woman_message, subject, message)

    time.sleep(1)
    gender.pop(0)
    firstname.pop(0)
    emails.pop(0)


    if not gender:
        break
=======

