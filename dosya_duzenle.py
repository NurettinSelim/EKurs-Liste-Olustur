import pickle

from openpyxl import load_workbook


# baslik = "DİNAR FEN LİSESİ 2019 - 2020 EĞİTİM ÖĞRETİM YILI {} SINIFI {} DERSİ 2. DÖNEM KURS YOKLAMA LİSTESİ"

def ogrenci_olustur(satir_no, numara, ad, biyoloji, cografya, fizik, kimya, matematik, tarih, edebiyat):
    ogrenci = dict()
    ogrenci["satir"] = satir_no
    ogrenci["numara"] = numara
    ogrenci["ad"] = ad
    if biyoloji == None:
        ogrenci["biyoloji"] = ""
        biyoloji = ""
    elif biyoloji == "" or biyoloji[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["biyoloji"] = ""
    else:
        if biyoloji[-2] == " ":
            if biyoloji[-4].isdigit():
                ogrenci["biyoloji"] = biyoloji[-4] + biyoloji[-3] + biyoloji[-1]
            else:
                ogrenci["biyoloji"] = biyoloji[-3] + biyoloji[-1]
        else:
            ogrenci["biyoloji"] = biyoloji[-5] + biyoloji[-4] + biyoloji[-2] + biyoloji[-1]

    if cografya == None:
        ogrenci["cografya"] = ""
        cografya = ""
    elif cografya == "" or cografya[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["cografya"] = ""
    else:
        if cografya[-2] == " ":
            if cografya[-4].isdigit():
                ogrenci["cografya"] = cografya[-4] + cografya[-3] + cografya[-1]
            else:
                ogrenci["cografya"] = cografya[-3] + cografya[-1]
        else:
            ogrenci["cografya"] = cografya[-5] + cografya[-4] + cografya[-2] + cografya[-1]

    if fizik == None:
        ogrenci["fizik"] = ""
        fizik = ""
    elif fizik == "" or fizik[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["fizik"] = ""
    else:
        if fizik[-2] == " ":
            if fizik[-4].isdigit():
                ogrenci["fizik"] = fizik[-4] + fizik[-3] + fizik[-1]
            else:
                ogrenci["fizik"] = fizik[-3] + fizik[-1]
        else:
            ogrenci["fizik"] = fizik[-5] + fizik[-4] + fizik[-2] + fizik[-1]

    if kimya == None:
        ogrenci["kimya"] = ""
        kimya = ""
    elif kimya == "" or kimya[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["kimya"] = ""
    else:
        if kimya[-2] == " ":
            if kimya[-4].isdigit():
                ogrenci["kimya"] = kimya[-4] + kimya[-3] + kimya[-1]
            else:
                ogrenci["kimya"] = kimya[-3] + kimya[-1]
        else:
            ogrenci["kimya"] = kimya[-5] + kimya[-4] + kimya[-2] + kimya[-1]

    if matematik == None:
        ogrenci["matematik"] = ""
        matematik = ""
    elif matematik == "" or matematik[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["matematik"] = ""
    else:
        if matematik[-2] == " ":
            if matematik[-4].isdigit():
                ogrenci["matematik"] = matematik[-4] + matematik[-3] + matematik[-1]
            else:
                ogrenci["matematik"] = matematik[-3] + matematik[-1]
        else:
            ogrenci["matematik"] = matematik[-5] + matematik[-4] + matematik[-2] + matematik[-1]

    if tarih == None:
        ogrenci["tarih"] = ""
        tarih = ""
    elif tarih == "" or tarih[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["tarih"] = ""
    else:
        if tarih[-2] == " ":
            if tarih[-4].isdigit():
                ogrenci["tarih"] = tarih[-4] + tarih[-3] + tarih[-1]
            else:
                ogrenci["tarih"] = tarih[-3] + tarih[-1]
        else:
            ogrenci["tarih"] = tarih[-5] + tarih[-4] + tarih[-2] + tarih[-1]

    if edebiyat == None:
        ogrenci["edebiyat"] = ""
        edebiyat = ""
    elif edebiyat == "" or edebiyat[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["edebiyat"] = ""
    else:
        if edebiyat[-2] == " ":
            if edebiyat[-4].isdigit():
                ogrenci["edebiyat"] = edebiyat[-4] + edebiyat[-3] + edebiyat[-1]
            else:
                ogrenci["edebiyat"] = edebiyat[-3] + edebiyat[-1]
        else:
            ogrenci["edebiyat"] = edebiyat[-5] + edebiyat[-4] + edebiyat[-2] + edebiyat[-1]

    ogrenci["biyoloji_ogr"] = biyoloji.split("\n")[0][8:]
    ogrenci["cografya_ogr"] = cografya.split("\n")[0][8:]
    ogrenci["fizik_ogr"] = fizik.split("\n")[0][8:]
    ogrenci["kimya_ogr"] = kimya.split("\n")[0][8:]
    ogrenci["matematik_ogr"] = matematik.split("\n")[0][8:]
    ogrenci["tarih_ogr"] = tarih.split("\n")[0][8:]
    ogrenci["edebiyat_ogr"] = edebiyat.split("\n")[0][8:]

    return ogrenci

numara_col = "A"
ad_col = "C"
okul_col = "D"
biyoloji_col = "F"
cografya_col = "G"
fizik_col = "H"
kimya_col = "I"
matematik_col = "J"
tarih_col = "K"
edebiyat_col = "L"

wb = load_workbook(filename='yeni_liste.xlsx')
ws = wb[wb.sheetnames[0]]

ogrenciler = list()

for row in range(3, ws.max_row + 1):
    # ws[cell_name].value
    okul_cell = "{}{}".format(okul_col, row)
    okul = ws[okul_cell].value
    okul = " "
    numara_cell = "{}{}".format(numara_col, row)
    numara = ws[numara_cell].value

    ad_cell = "{}{}".format(ad_col, row)
    ad = ws[ad_cell].value

    biyoloji_cell = "{}{}".format(biyoloji_col, row)
    biyoloji = ws[biyoloji_cell].value

    cografya_cell = "{}{}".format(cografya_col, row)
    cografya = ws[cografya_cell].value

    fizik_cell = "{}{}".format(fizik_col, row)
    fizik = ws[fizik_cell].value

    kimya_cell = "{}{}".format(kimya_col, row)
    kimya = ws[kimya_cell].value

    matematik_cell = "{}{}".format(matematik_col, row)
    matematik = ws[matematik_cell].value

    tarih_cell = "{}{}".format(tarih_col, row)
    tarih = ws[tarih_cell].value

    edebiyat_cell = "{}{}".format(edebiyat_col, row)
    edebiyat = ws[edebiyat_cell].value
    if edebiyat == None:
        edebiyat = ""

    ogrenciler.append(
        ogrenci_olustur(row, numara, ad, biyoloji, cografya, fizik, kimya, matematik, tarih, edebiyat=edebiyat))

ogrenciler = sorted(ogrenciler, key=lambda i: i['numara'])
olustur = list()
for ogrenci in ogrenciler:
    if ogrenci["biyoloji"]:
        if "biyoloji" + ogrenci["biyoloji"] not in olustur:
            olustur.append("biyoloji" + ogrenci["biyoloji"])
    if ogrenci["cografya"]:
        if "cografya" + ogrenci["cografya"] not in olustur:
            olustur.append("cografya" + ogrenci["cografya"])
    if ogrenci["fizik"]:
        if "fizik" + ogrenci["fizik"] not in olustur:
            olustur.append("fizik" + ogrenci["fizik"])
    if ogrenci["fizik"]:
        if "fizik" + ogrenci["fizik"] not in olustur:
            olustur.append("fizik" + ogrenci["fizik"])
    if ogrenci["kimya"]:
        if "kimya" + ogrenci["kimya"] not in olustur:
            olustur.append("kimya" + ogrenci["kimya"])
    if ogrenci["matematik"]:
        if "matematik" + ogrenci["matematik"] not in olustur:
            olustur.append("matematik" + ogrenci["matematik"])
    if ogrenci["tarih"]:
        if "tarih" + ogrenci["tarih"] not in olustur:
            olustur.append("tarih" + ogrenci["tarih"])
    if ogrenci["edebiyat"]:
        if "edebiyat" + ogrenci["edebiyat"] not in olustur:
            olustur.append("edebiyat" + ogrenci["edebiyat"])


# DOSYA OLUŞTUR
# kurs_wb = pickle.load(open('lib.dll', 'rb'))
kurs_wb = load_workbook('sablon.xlsx')

for i in range(len(olustur) - 1):
    kurs_wb.copy_worksheet(kurs_wb.active)

i = 0
for sheet in kurs_wb:
    sheet.title = olustur[i]
    i += 1

import time

ay = int(time.strftime("%m"))
yil = time.strftime("%Y")
if ay > 8:
    donem = yil + str(int(yil) + 1)
else:
    donem = str(int(yil) - 1) + " - " + yil

if 2 <= ay <= 6:
    donem_2 = "2. DÖNEM"
elif ay == 1 or 9 >= ay:
    donem_2 = "1. DÖNEM"
elif 7 == ay or 8 == ay:
    donem_2 = "YAZ DÖNEMİ"
else:
    donem_2 = " "

baslik = okul + " " + donem + " EĞİTİM ÖĞRETİM YILI {} SINIFI {} DERSİ " + donem_2 + " KURS YOKLAMA LİSTESİ"

for sheet in kurs_wb:
    for k in range(len(sheet.title)):
        if sheet.title[k].isdigit():
            ders = sheet.title[:k]
            break
    k = 5
    for ogrenci in ogrenciler:
        k = 5
        if ogrenci[ders]:
            if ders + ogrenci[ders] == sheet.title:
                sheet["D2"].value = ogrenci[ders + "_ogr"]
                sheet["A1"].value = baslik.format(ogrenci[ders], ders.upper())
                if sheet["B" + str(k)].value is None:
                    sheet["B" + str(k)].value = ogrenci["numara"]
                    sheet["C" + str(k)].value = ogrenci["ad"]
                else:
                    while sheet["B" + str(k)].value is not None:
                        k += 1
                    sheet["B" + str(k)].value = ogrenci["numara"]
                    sheet["C" + str(k)].value = ogrenci["ad"]

kurs_wb.save('listeler.xlsx')
