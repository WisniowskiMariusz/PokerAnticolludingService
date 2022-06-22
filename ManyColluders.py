import glob

def GetSheet(*args):    #Get data about spreadsheet
    # get the doc from the scripting context which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    # get the XText interface
    sheet = model.Sheets.getByIndex(0)
    return sheet

def GetPath(*args):    # modify path to directory
    sheet = GetSheet()

    if  sheet.getCellByPosition(5, 1).String == "":
        sheet.getCellByPosition(5, 1).String=("Podaj katalog!")
        return
    else:
        path = sheet.getCellByPosition(5, 1).String
        path = path.replace("\\","\\\\") + "\\\\"
    return path

def ClearTable(*args):  #Clear results
    sheet = GetSheet()
    max = int(sheet.getCellByPosition(1, 1).Value)
    for l in range(max):
        sheet.getCellByPosition(0, l + 3).String =""
        sheet.getCellByPosition(1, l + 3).String =""
        sheet.getCellByPosition(2, l + 3).String =""
        sheet.getCellByPosition(3, l + 3).String =""
        sheet.getCellByPosition(4, l + 3).String =""
        sheet.getCellByPosition(5, l + 3).String =""


def PrintOpponents(Oppo,sheet,hands,sessions,TotalOppo,y,s):  #Prints results
    sheet.getCellByPosition(0, y).String = "Total"
    sheet.getCellByPosition(1, y).Value = hands
    sheet.getCellByPosition(2, y).String = s
    sheet.getCellByPosition(3, y).Value = sessions
    sheet.getCellByPosition(4, y).String = "Opponents"
    sheet.getCellByPosition(5, y).Value = TotalOppo
    sheet.getCellByPosition(0, y + 1).String = "ID"
    sheet.getCellByPosition(1, y + 1).String = "Hands"
    sheet.getCellByPosition(2, y + 1).String = "Hands%"
    sheet.getCellByPosition(3, y + 1).String = "Sessions"
    sheet.getCellByPosition(4, y + 1).String = "Sessions%"
    sheet.getCellByPosition(5, y + 1).String = "H2N note"
    for l in range(len(Oppo)):
        sheet.getCellByPosition(0, l + y + 2).Value = Oppo[l][0]
        sheet.getCellByPosition(1, l + y + 2).Value = Oppo[l][1]
        sheet.getCellByPosition(2, l + y + 2).Value = Oppo[l][1]/hands
        sheet.getCellByPosition(3, l + y + 2).Value = Oppo[l][2]
        sheet.getCellByPosition(4, l + y + 2).Value = Oppo[l][2]/sessions
        sheet.getCellByPosition(5, l + y + 2).String = (sheet.getCellByPosition(0, l + y + 2).String + " / " + sheet.getCellByPosition(2, l + y + 2).String + " / " + sheet.getCellByPosition(4, l + y + 2).String + " / " + sheet.getCellByPosition(1, y).String).strip()

def CheckIfInTheList(ID,Oppo):
# sprawdza czy ID jest już na liście Oppo i jeżeli jest to zwiększa licznik rozdań, dodatkowo jeżeli zwiększenie nastąpiło pierwszy raz to zwieksza też licznik sesji.
#Jeżeli elementu nie ma na liście to wstawia go na koniec.
    flaga = 0
    for o in range(len(Oppo)):
        if Oppo[o][0] == ID:
            flaga = 1
            Oppo[o][1] += 1
            if Oppo[o][3] == 0:
                Oppo[o][2] += 1
                Oppo[o][3] = 1
            break
    if flaga == 0:
        Oppo.append([ID, 1, 1, 1])  # Wstawia kolejny element


def Colluders(*args):
    # Path to tree
    path = GetPath()
    sheet = GetSheet()
    lastOppo = int(sheet.getCellByPosition(1, 1).Value)
    LastWritten= int(sheet.getCellByPosition(2, 1).Value)
    dircounter = 0
    directories = [d for d in glob.glob(path + "*\\")]
    for d in directories:
        dircounter += 1
        files = [f for f in glob.glob(d + "**/*.txt", recursive=True)]
        Oppo = []
        hands=0
        sessions=0
        for f in files:
            with open(f) as plik:
                sessions+=1 #licznik sesji
                i = 0
                for m in range(len(Oppo)): #zerowanie flag dla sesji
                    Oppo[m][3]=0
                for line in plik.readlines():
                    i += 1  # licznik linii
                    if f.count("PMS_")==1:
                        if line.strip().count("Game started at:") == 1:
                            hands += 1  # licznik rozdań
                        elif line.strip().count("ante") == 1:
                            ID = line.strip().split()[1].replace(":", "")
                            CheckIfInTheList(ID, Oppo)
                    else:
                        if line.strip().count("PokerMaster Hand #") == 1:
                            hands += 1  # licznik rozdań
                        elif line.strip().count("ante") == 1:
                            ID = line.strip().split()[0].replace(":", "")
                            CheckIfInTheList(ID, Oppo)
        if sheet.getCellByPosition(1, 1).Value > len(Oppo):
            lastOppo = len(Oppo)
        if LastWritten < 1:
            LastWritten = 1
        Opposessions = sorted(Oppo, reverse=True, key=lambda list: (list[2],list[1]))[0:lastOppo]
        Oppohands = sorted(Oppo, reverse=True, key=lambda list: (list[1],list[2]))[0:lastOppo]
        if sheet.getCellByPosition(3, 1).String=="h":
            PrintOpponents(Oppohands,sheet,hands,sessions,len(Oppo),LastWritten+2,"by Hands")
            PrintOpponents(Opposessions,sheet,hands,sessions,len(Oppo),lastOppo+5+LastWritten,"by Sessions")
        else:
            PrintOpponents(Opposessions,sheet,hands,sessions,len(Oppo),LastWritten+2,"by Sessions")
            PrintOpponents(Oppohands,sheet,hands,sessions,len(Oppo),lastOppo+5+LastWritten,"by Hands")
        LastWritten = lastOppo*2 + LastWritten + 6
    sheet.getCellByPosition(2, 1).Value = LastWritten
    sheet.getCellByPosition(4, 1).Value = sheet.getCellByPosition(4, 1).Value + dircounter
    sheet.getCellByPosition(4, 2).Value = dircounter