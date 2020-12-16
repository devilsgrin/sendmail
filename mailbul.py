"""
Author: Emir Polat
Date: 21.11.2020
"""

from itertools import chain
import xlrd

def ExtractNames():
    file = (your_xlsx_file)

    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    extension = input("Mail Domaini: ")

    mails = []
    # Listedeki bütün her şeyi getirerek bir listeye atamak için yapılan döngü
    for i in range(sheet.nrows):
        name = sheet.row_values(i)
        all_surname = name[1]
        all_first_char = list(chain(name[0]))
        all_mails = all_first_char[0] + all_surname + "@" + extension

        mails.append(all_mails)

    with open("new_mails.txt", "w") as f:
        for item in mails:
            f.write("%s\n" % item)

        print("Bütün mailler şuraya yazıldı: new_mails.txt")
ExtractNames()
