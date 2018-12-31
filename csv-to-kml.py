import simplekml
from openpyxl import load_workbook
import re
import codecs

HoL = {}

with codecs.open('whc-en.kml','r',encoding='utf-8', errors='ignore') as fwhc:

    buffer = ''
    drinnen = False

    for line in fwhc:
        if re.search(r"<description>",line):
            drinnen = True
            continue
        if re.search(r"<\/description>",line):
            drinnen = False
        #    m = re.search(r"site_([0-9]+)\.jpg",buffer)
            m = re.search(r"org\/en\/list\/([0-9]+)\"",buffer)
            if bool(m):
                id= m.groups()
                #print (id[0])
                HoL[str(id[0])] = buffer
            else:
                print(buffer)
            buffer = ''
            continue
        if drinnen:
            buffer = buffer + line


print("Laenge vom Hash")
print (len(HoL))

for key in HoL:
    print(key)

ID_NO = 0
NAME_EN = 1
SHORT_DESCRIPTION_EN = 2
JUSTIFICATION_EN = 3
LONGITUDE = 4
LATITUDE = 5
CATEGORY = 6
PICTURE = 7


kml = simplekml.Kml()


kml.document.name = "Unesco World Heritage Sites"

folCult = kml.newfolder(name='Culture')
folNatu = kml.newfolder(name='Nature')
folMix = kml.newfolder(name='Mix')


wb = load_workbook(filename = 'welterbestaetten.xlsx')
ws = wb.active

#print(HoL['1009'])



for row in ws.values:
    if row[ID_NO] == "id_no":
        continue
    elif row[CATEGORY] == "Cultural":
        pnt = folCult.newpoint()
    elif row[CATEGORY] == "Natural":
        pnt = folNatu.newpoint()
    elif row[CATEGORY] == "Mixed":
        pnt = folMix.newpoint()
    else:
        continue
    pnt.name = row[NAME_EN]
    pnt.coords = [(row[LONGITUDE],row[LATITUDE])]
    #print(row[ID_NO])
    if str(row[ID_NO]) in HoL.keys():
        pnt.description = HoL[str(row[ID_NO])]
    else:
        pnt.description = row[SHORT_DESCRIPTION_EN]

kml.savekmz("temp.kmz")
