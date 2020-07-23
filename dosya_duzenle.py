from openpyxl import load_workbook
import pickle
import term
from models import Student

# baslik = "DİNAR FEN LİSESİ 2019 - 2020 EĞİTİM ÖĞRETİM YILI {} SINIFI {} DERSİ 2. DÖNEM KURS YOKLAMA LİSTESİ"

def ogrenci_olustur(satir_no, numara, ad, biyoloji, cografya, fizik, kimya, matematik, tarih, edebiyat):
    ogrenci = dict()
    ogrenci["satir"] = satir_no
    ogrenci["numara"] = numara
    ogrenci["ad"] = ad

    if biyoloji is None or biyoloji == "" or biyoloji[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["biyoloji"] = ""
        biyoloji = ""
    else:
        if biyoloji[-2] == " ":
            if biyoloji[-4].isdigit():
                ogrenci["biyoloji"] = biyoloji[-4] + biyoloji[-3] + biyoloji[-1]
            else:
                ogrenci["biyoloji"] = biyoloji[-3] + biyoloji[-1]
        elif biyoloji[-3] == " ":
            ogrenci["biyoloji"] = biyoloji[-5] + biyoloji[-4] + biyoloji[-2] + biyoloji[-1]

    if cografya is None or cografya == "" or cografya[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["cografya"] = ""
        cografya = ""
    else:
        if cografya[-2] == " ":
            if cografya[-4].isdigit():
                ogrenci["cografya"] = cografya[-4] + cografya[-3] + cografya[-1]
            else:
                ogrenci["cografya"] = cografya[-3] + cografya[-1]
        elif cografya[-3] == " ":
            ogrenci["cografya"] = cografya[-5] + cografya[-4] + cografya[-2] + cografya[-1]

    if fizik is None or fizik == "" or fizik[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["fizik"] = ""
        fizik = ""
    else:
        if fizik[-2] == " ":
            if fizik[-4].isdigit():
                ogrenci["fizik"] = fizik[-4] + fizik[-3] + fizik[-1]
            else:
                ogrenci["fizik"] = fizik[-3] + fizik[-1]
        elif fizik[-3] == " ":
            ogrenci["fizik"] = fizik[-5] + fizik[-4] + fizik[-2] + fizik[-1]

    if kimya is None or kimya == "" or kimya[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["kimya"] = ""
        kimya = ""
    else:
        if kimya[-2] == " ":
            if kimya[-4].isdigit():
                ogrenci["kimya"] = kimya[-4] + kimya[-3] + kimya[-1]
            else:
                ogrenci["kimya"] = kimya[-3] + kimya[-1]
        elif kimya[-3] == " ":
            ogrenci["kimya"] = kimya[-5] + kimya[-4] + kimya[-2] + kimya[-1]

    if matematik is None or matematik == "" or matematik[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["matematik"] = ""
        matematik = ""
    else:
        if matematik[-2] == " ":
            if matematik[-4].isdigit():
                ogrenci["matematik"] = matematik[-4] + matematik[-3] + matematik[-1]
            else:
                ogrenci["matematik"] = matematik[-3] + matematik[-1]
        elif matematik[-3] == " ":
            ogrenci["matematik"] = matematik[-5] + matematik[-4] + matematik[-2] + matematik[-1]

    if tarih is None or tarih == "" or tarih[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["tarih"] = ""
        tarih = ""
    else:
        if tarih[-2] == " ":
            if tarih[-4].isdigit():
                ogrenci["tarih"] = tarih[-4] + tarih[-3] + tarih[-1]
            else:
                ogrenci["tarih"] = tarih[-3] + tarih[-1]
        elif tarih[-3] == " ":
            ogrenci["tarih"] = tarih[-5] + tarih[-4] + tarih[-2] + tarih[-1]

    if edebiyat is None or edebiyat == "" or edebiyat[-18:] == "(Kursa Yerleşmedi)":
        ogrenci["edebiyat"] = ""
        edebiyat = ""
    else:
        if edebiyat[-2] == " ":
            if edebiyat[-4].isdigit():
                ogrenci["edebiyat"] = edebiyat[-4] + edebiyat[-3] + edebiyat[-1]
            else:
                ogrenci["edebiyat"] = edebiyat[-3] + edebiyat[-1]
        elif edebiyat[-3] == " ":
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

excel_ogr_list = "excels/yeni_liste.xlsx"
wb = load_workbook(filename=excel_ogr_list)
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

    ogrenciler.append(
        Student(row, numara, ad, biyoloji, cografya, fizik, kimya, matematik, tarih, edebiyat))

ogrenciler = sorted(ogrenciler, key=lambda i: i.numara)
olustur_list = list()

for ogrenci in ogrenciler:
    if ogrenci.biyoloji.name:
        olustur_list.append("biyoloji" + ogrenci.biyoloji.name)
    if ogrenci.cografya.name:
        olustur_list.append("cografya" + ogrenci.cografya.name)
    if ogrenci.fizik.name:
        olustur_list.append("fizik" + ogrenci.fizik.name)
    if ogrenci.kimya.name:
        olustur_list.append("kimya" + ogrenci.kimya.name)
    if ogrenci.matematik.name:
        olustur_list.append("matematik" + ogrenci.matematik.name)
    if ogrenci.tarih.name:
        olustur_list.append("tarih" + ogrenci.tarih.name)
    if ogrenci.edebiyat.name:
        olustur_list.append("edebiyat" + ogrenci.edebiyat.name)

olustur_set = set(olustur_list)
olustur_list = list(olustur_set)

# DOSYA OLUŞTUR
kurs_wb = pickle.load(open('lib.dll', 'rb'))
# kurs_wb = load_workbook('sablon.xlsx')

for i in range(len(olustur_list) - 1):
    kurs_wb.copy_worksheet(kurs_wb.active)

for number, sheet in enumerate(kurs_wb):
    sheet.title = olustur_list[number]

baslik = "{} {} EĞİTİM ÖĞRETİM YILI {} SINIFI {} DERSİ {} KURS YOKLAMA LİSTESİ"

for sheet in kurs_wb:
    for k in range(len(sheet.title)):
        if sheet.title[k].isdigit():
            ders = sheet.title[:k]
            break
    for ogrenci in ogrenciler:
        k = 5
        if ders + ogrenci.lesson_from_name(ders).name == sheet.title:
            sheet["D2"].value = ogrenci.lesson_from_name(ders).teacher
            sheet["A1"].value = baslik.format(okul,
                                              term.term_year(),
                                              ogrenci.lesson_from_name(ders).name,
                                              ders.upper(),
                                              term.term_year())
            if sheet["B" + str(k)].value is None:
                sheet["B" + str(k)].value = ogrenci.numara
                sheet["C" + str(k)].value = ogrenci.ad
            else:
                while sheet["B" + str(k)].value is not None:
                    k += 1
                sheet["B" + str(k)].value = ogrenci.numara
                sheet["C" + str(k)].value = ogrenci.ad

kurs_wb.save('excels/listeler.xlsx')
