import time

ay = int(time.strftime("%m"))
yil = time.strftime("%Y")
if ay > 8:
    donem = yil + str(int(yil) + 1)
else:
    donem = str(int(yil) - 1) + " - " + yil

if 7 == ay or 8 == ay:
    donem_2 = "YAZ DÖNEMİ"
elif 2 <= ay <= 6:
    donem_2 = "2. DÖNEM"
elif ay == 1 or 9 >= ay:
    donem_2 = "1. DÖNEM"
else:
    donem_2 = " "

print(donem, donem_2)
