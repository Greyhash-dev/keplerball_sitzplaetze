import pandas as pd
import sys
df = pd.read_excel('Klassenliste.xlsx')

ver = ['K1', 'K2', 'K3', 'K4', 'K5', 'K6', 'K7', 'K8', 'K9', 'K10', 'K11', 'K12', 'K13', 'K14', 'K15', 'K16', 'K17', 'K18', 'K19', 'K20', 'K21', 'K22', 'K23', 'K24', 'K25', 'K26', 'K27', 'K28', 'K29', 'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'L10', 'L11', 'L12', 'L13', 'L14', 'L15', 'L16', 'L17', 'L18', 'L19', 'L20', 'L21', 'L22', 'L23', 'L24', 'L25']

personen = []
for col in df.iterrows():
    if str(col[1][0]) != "nan":
        klasse = col[1][0]
    if col[0] > 82:
        break
    if col[1]["Punkte am Ball"] in ["kommt net", "-"]:
        continue
    sitz = ""
    if str(col[1]["Sitzplatz"]) != "nan":
        sitz = str(col[1]["Sitzplatz"])
    personen.append({"Name":str(col[1]["Vorname"]) + " " + str(col[1]["Nachname"]), "Key":col[0], "Punkte":col[1]["Punkte am Ball"], "Klasse":klasse, "Sitzplatz":sitz})
print("Personen: "+ str(len(personen)))

bes = 0
for platz in df["Sitzplatz"]:
    if str(platz) != "nan":
        for a in str(platz).split(";"):
            ver.remove(a)
            bes += 1

personen = sorted(personen, key=lambda x: x["Punkte"], reverse=False)
last_punkte = -1000
for pers in personen:
    if last_punkte != pers["Punkte"]:
        print("="*15 + " " + str(pers["Punkte"]) + " " + "="*15)
        print("")
        last_punkte = pers["Punkte"]
    print(str(pers["Key"]) + " " + pers["Name"] + ":")
    print("\tKlasse: " + pers["Klasse"] + " | " + pers["Sitzplatz"])
    print("")


print("== Verfügbare Sitzplätze ==")
last = ver[0]
was = 0
for i in range(0, len(ver)-1):
    if int(last[1:]) + 2 == int(ver[i+1][1:]):
        last = ver[i]
        print(".", end="")
        was = 0
    else:
        if was == 1:
            print(", ",end="")
        was = 1
        print(ver[i], end="")
        last = ver[i]
print(ver[-1])
print("Besetzte Sitzplätze: "+str(bes))
print("Verfügbar: "+str(len(ver)))
