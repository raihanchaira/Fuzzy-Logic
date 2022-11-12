import pandas as pd
import xlsxwriter

# memasukkan tabel doorsmeer
dataBengkelMobil = pd.read_csv("bengkel.csv")
#print(dataBengkelMobil)

# mendefinisikan kolom pada tabel
id = dataBengkelMobil["id"]
kualitas = dataBengkelMobil["servis"]
harga = dataBengkelMobil["harga"]

#array untuk menyimpan hasil defuzzy
arrDefuzzy = []
#array untuk menyimpan id
arrid = []

""""
nilai linguistik :
    sangat bagus  0
    bagus         1   
    sedang        2
    jelek         3   
    sangat jelek  4

    hargaMurah    0
    hargaSedang   1
    hargaMahal    2
"""


#PROSES FUZZYFICATION
def fuzzyfication(kualitas, harga):
    #kualitas
    global kualitasSangatJelek, kualitasJelek, kualitasSedang, kualitasBagus, kualitasSangatBagus
    kualitasSangatJelek = 0
    kualitasJelek = 0
    kualitasSedang = 0
    kualitasBagus = 0
    kualitasSangatBagus = 0
    if kualitas <= 15:
        kualitasSangatJelek = 1
        kualitasArr[4] = kualitasSangatJelek
        SangatJelek = kualitasSangatJelek
    elif kualitas > 15 and kualitas < 20:
        kualitasSangatJelek = (20 - kualitas) / (20 - 15)
        kualitasArr[4] = kualitasSangatJelek
        kualitasJelek = (kualitas - 15) / (20 - 15)
        kualitasArr[3] = kualitasJelek
    elif kualitas >= 20 and kualitas <= 35:
        kualitasJelek = 1
        kualitasArr[3] = kualitasJelek
    elif kualitas > 35 and kualitas < 40:
        kualitasJelek = (40 - kualitas) / (40 - 35)
        kualitasArr[3] = kualitasJelek
        kualitasSedang = (kualitas - 35) / (40 - 35)
        kualitasArr[2] = kualitasSedang
    elif kualitas >= 40 and kualitas <= 55:
        kualitasSedang = 1
        kualitasArr[2] = kualitasSedang
    elif kualitas > 55 and kualitas < 60:
        kualitasSedang = (60 - kualitas) / (60 - 55)
        kualitasArr[2] = kualitasSedang
        kualitasBagus = (kualitas - 55) / (60 - 55)
        kualitasArr[1] = kualitasBagus
    elif kualitas >= 60 and kualitas <= 75:
        kualitasBagus = 1
        kualitasArr[1] = kualitasBagus
    elif kualitas > 75 and kualitas < 80:
        kualitasBagus = (80 - kualitas) / (80 - 75)
        kualitasArr[1] = kualitasBagus
        kualitasSangatBagus = (kualitas - 75) / (80 - 75)
        kualitasArr[0] = kualitasSangatBagus
    elif kualitas >= 80 and kualitas <= 100:
        kualitasSangatBagus = 1
        kualitasArr[0] = kualitasSangatBagus

    # harga
    global hargaMurah, hargaSedang, hargaMahal
    hargaMurah = 0
    hargaSedang = 0
    hargaMahal = 0
    if harga <= 4:
        hargaMurah = 1
        hargaArr[0] = hargaMurah
    elif harga > 4 and harga < 5:
        hargaMurah = (5 - harga) / (5 - 4)
        hargaArr[0] = hargaMurah
    elif harga > 4 and harga < 6:
        hargaSedang = (harga - 4) / (6 - 4)
        hargaArr[1] = hargaSedang
    elif harga == 6:
        hargaSedang = 1
        hargaArr[1] = hargaSedang
    elif harga > 6 and harga < 8:
        hargaSedang = (8 - harga) / (8 - 6)
        hargaArr[1] = hargaSedang
    elif harga > 7 and harga < 8:
        hargaSedang = (harga - 7) / (8 - 7)
        hargaArr[1] = hargaSedang
    elif harga >= 8:
        hargaMahal = 1
        hargaArr[2] = hargaMahal

    return kualitasArr, hargaArr


for i in range(100):
    kualitasArr = [0, 0, 0, 0, 0]  # [sangat bagus, bagus, sedang, jelek, sangat jelek]
    hargaArr = [0, 0, 0]  # [hargaMurah, hargaSedang, hargaMahal]
    hasil1, hasil2 = fuzzyfication(kualitas[i], harga[i])
    arrid.append(i+1)

    #PROSES INFERENCE
    """"
    not recomend = NR
    recomend = R
    very recomend = VR
    """
    VR = []
    if hargaArr[0] == hargaMurah and kualitasArr[0] == kualitasSangatBagus:
        VR.append(min(hargaArr[0], kualitasArr[0]))
    if hargaArr[0] == hargaMurah and kualitasArr[1] == kualitasBagus:
        VR.append(min(hargaArr[0], kualitasArr[1]))
    if hargaArr[1] == hargaSedang and kualitasArr[0] == kualitasSangatBagus:
        VR.append(min(hargaArr[1], kualitasArr[0]))
    if hargaArr[1] == hargaSedang and kualitasArr[1] == kualitasBagus:
        VR.append(min(hargaArr[1], kualitasArr[1]))

    R = []
    if hargaArr[0] == hargaMurah and kualitasArr[2] == kualitasSedang:
        R.append(min(hargaArr[0], kualitasArr[2]))
    if hargaArr[1] == hargaSedang and kualitasArr[2] == kualitasSedang:
        R.append(min(hargaArr[1], kualitasArr[2]))
    if hargaArr[2] == hargaMahal and kualitasArr[0] == kualitasSangatBagus:
        R.append(min(hargaArr[2], kualitasArr[0]))
    if hargaArr[2] == hargaMahal and kualitasArr[1] == kualitasBagus:
        R.append(min(hargaArr[2], kualitasArr[1]))

    NR = []
    if hargaArr[2] == hargaMahal and kualitasArr[2] == kualitasSedang:
        NR.append(min(hargaArr[2], kualitasArr[2]))
    if hargaArr[0] == hargaMurah and kualitasArr[3] == kualitasJelek:
        NR.append(min(hargaArr[0], kualitasArr[3]))
    if hargaArr[0] == hargaMurah and kualitasArr[4] == kualitasSangatJelek:
        NR.append(min(hargaArr[0], kualitasArr[4]))
    if hargaArr[1] == hargaSedang and kualitasArr[3] == kualitasJelek:
        NR.append(min(hargaArr[1], kualitasArr[3]))
    if hargaArr[1] == hargaSedang and kualitasArr[4] == kualitasSangatJelek:
        NR.append(min(hargaArr[1], kualitasArr[4]))
    if hargaArr[2] == hargaMahal and kualitasArr[3] == kualitasJelek:
        NR.append(min(hargaArr[2], kualitasArr[3]))
    if hargaArr[2] == hargaMahal and kualitasArr[4] == kualitasSangatJelek:
        NR.append(min(hargaArr[2], kualitasArr[4]))

    vr = max(VR)
    r = max(R)
    nr = max(NR)

    #PROSES DEFUZZYFICATION
    pengali = (vr*100)+(r*70)+(nr*50)
    pembagi = vr + r + nr
    hasil = pengali/pembagi
    arrDefuzzy.append(hasil)


#PROSES OUTPUT (XLSX)
dataBengkelMobil["score"] = arrDefuzzy
sorting = dataBengkelMobil.sort_values(["score"], ascending=False)
hasil = sorting.head(10)

workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'id')
worksheet.write('B1', 'score')

row = 1
column = 0

for item in hasil.id:
    worksheet.write(row, column, item)
    row += 1

row = 1
for item in hasil.score:
    worksheet.write(row, 1, item)
    row += 1

workbook.close()

