# Variabeln
## Pfad zum gewünschten Excel Dokument
path_to_file = "Liste KNX def 16_3_22.xlsx"

from openpyxl import load_workbook

# Hauptfuntion die (momentan) das Dokument in eine Python Datenstruktur bringt
def main():
    file, sheet, = import_file(path_to_file)
    data = get_data(sheet)
    display_all(data)



# Testfunktion die alle Geschosse mit allen Aktionen und Messwerten anzeigt
def display_all(data):
    for geschoss in data:
        print()
        print("----------------------------------------------")
        print("Geschoss: {}".format(geschoss["name"]))
        for aktion in geschoss["aktion"]:
            print("----------------------------------------------")
            print("\nAktion: {}\n".format(aktion["name"]))
            for messwert in aktion["messwert"]:
                print("Messwert: {}".format(messwert["name"]))
                #print("aufschalten: {}".format(messwert["aufschalten"]))
        print("----------------------------------------------")

# Importiert das Dokument und gibt dieses sowie das aktuelle Arbeitsblatt zurück
def import_file(path):
    wb = load_workbook(filename = path) 
    ws = wb.active
    return wb, ws 

# Wandelt das Arbeitsblatt in eine Python Datenstruktur bringt
# data: Liste mit Geschossen
# geschoss: Dictionary mit Atributen und Liste von Aktionen
# aktion: Dictionary mit Atributen und Liste von Messwerten
# messwert: Dictionary mit Atributen
def get_data(sheet):
    data = []
    for row in sheet.iter_rows(min_col = 5, max_col = 5, max_row = 2000):
        for cell in row:
            if cell.value == None:
                continue
            split = str(cell.value).split("/")
            if split[1] == "-" and split[2] == "-":
                name = sheet.cell(row[0].row, column = 1).value
                geschoss = {"name": name, 
                            "id_geschoss": int(split[0]), 
                            "aktion": []}
                data.append(geschoss)
            elif split[2] == "-":
                name = sheet.cell(row[0].row, column = 2).value
                bool_f = sheet.cell(row[0].row, column=6).value == "true"
                aktion = {"name": name, 
                          "id_aktion": int(split[1]), 
                          "bool_f": bool_f,
                          "messwert": []}
                data[-1]["aktion"].append(aktion)
            else:
                name = sheet.cell(row[0].row, column = 3).value
                aufschalten = sheet.cell(row[0].row, column=4).value == "x"
                bool_f = sheet.cell(row[0].row, column=6).value == "true"
                kommentar = sheet.cell(row[0].row, column = 7).value
                dpst = sheet.cell(row[0].row, column = 8).value
                messwert = {"name": name,
                            "aufschalten": aufschalten,
                            "id_messwert": int(split[2]),
                            "bool_f": bool_f,
                            "kommentar": kommentar,
                            "dpst": dpst, }               
                data[-1]["aktion"][-1]["messwert"].append(messwert)
    return data

if __name__ == '__main__':
  main()
