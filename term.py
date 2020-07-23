import time


def term_year():
    ay = int(time.strftime("%m"))
    yil = time.strftime("%Y")
    if ay > 8:
        term_year = yil + str(int(yil) + 1)
    else:
        term_year = str(int(yil) - 1) + " - " + yil

    return term_year


def term_name():
    ay = int(time.strftime("%m"))

    if 7 == ay or 8 == ay:
        term_name = "YAZ DÖNEMİ"
    elif 2 <= ay <= 6:
        term_name = "2. DÖNEM"
    elif ay == 1 or 9 >= ay:
        term_name = "1. DÖNEM"
    else:
        term_name = " "

    return term_name
