import datetime
import xlwings as xw
import os
from lib import Information
from pathlib import Path


# GlobalVar: current excel file
datei_name_excel = ""
_ergebnis_ordner = ""
aktuelle_zeile = 0


# noinspection PyBroadException
def erstelle_excel_datei(ergebnis_ordner):

    """
    Erstelle neue Excel Datei.\n
    Gebe ihr einen Zeitstempel Namen.\n
    Schreibe den AnfangsText.\n
    """

    # Falls UnterOrdnder nicht existiert create ihn
    try:
        os.mkdir(ergebnis_ordner)
    except:
        pass

    # Merken des ErgebnisOrdners
    global _ergebnis_ordner
    _ergebnis_ordner = ergebnis_ordner

    # Konstruiere Namen - DayBased
    global datei_name_excel
    datei_name_excel = datetime.datetime.now().date().__str__()
    datei_name_excel = "Tickets_{}".format(datei_name_excel)

    # Checke ob Datei bereits existiert
    my_file = Path("{}/{}.xlsx".format(ergebnis_ordner, datei_name_excel))

    if not my_file.is_file():
        # Erzeuge neue Excel Datei
        wb = xw.Book()
        # Erzeuge Page
        ws = wb.sheets.add(datei_name_excel)
        # Schreibe Spalten Überschriften
        ws.range("A1").value = "AuftragsNr"
        ws.range("B1").value = "Postfach"
        ws.range("C1").value = "Datum"
        ws.range("D1").value = "Uhrzeit"
        ws.range("E1").value = "BA ID"
        ws.range("F1").value = "Beanstandung"
        ws.range("G1").value = "Fahrzeugtyp"
        ws.range("H1").value = "Km Stand"
        ws.range("I1").value = "MKB"
        ws.range("J1").value = "GKB"
        ws.range("K1").value = "Kundencodierung"
        ws.range("L1").value = "Werkstattcodierung"
        ws.range("M1").value = "Vz Nr"
        ws.range("N1").value = "B Nr"
        ws.range("O1").value = "Ansprechpartner"
        ws.range("P1").value = "Tel"
        ws.range("Q1").value = "Postfach Filter"
        ws.range("R1").value = "DissURL"
        # Fett & Autofit
        ws.range("A1", "R1").api.Font.Bold = True
        ws.autofit()
        # Speichere Excel Datei
        wb.save("{}/{}.xlsx".format(ergebnis_ordner, datei_name_excel))
        # Schließe Fenster
        # Schließe das Fenster
        app = xw.apps.active
        app.quit()
        global aktuelle_zeile
        aktuelle_zeile = 2
    else:
        # Finde raus welche Zeile gebraucht wird
        wb = xw.Book("{}/{}.xlsx".format(_ergebnis_ordner, datei_name_excel))
        # Öffne erste Page
        ws = wb.sheets[0]
        # Anzahl Rows
        aktuelle_zeile = ws.range('A1').current_region.last_cell.row + 1

        # Schließe das Fenster
        app = xw.apps.active
        app.quit()


def schreibe_zeile_in_excel(information: Information):
    """
    Schreibt eine Trigger Zeile in die Excel Datei für den aktuellen Run.

    :param information:
        Erwartet wird ein gefülltes Informationsobjekt. Dieses muss alle gelesenen Inhalte enthalten
    """

    # Öffne Excel Datei
    global _ergebnis_ordner
    wb = xw.Book("{}/{}.xlsx".format(_ergebnis_ordner, datei_name_excel))
    # Öffne erste Page
    ws = wb.sheets[0]
    # Schreibe in x-ter Zeile
    global aktuelle_zeile
    ws.range("A{}".format(aktuelle_zeile)).value = information.auftragsnummer
    ws.range("B{}".format(aktuelle_zeile)).value = information.postfach
    ws.range("C{}".format(aktuelle_zeile)).value = information.datum
    ws.range("D{}".format(aktuelle_zeile)).value = information.zeit
    ws.range("E{}".format(aktuelle_zeile)).value = information.ba_id
    ws.range("F{}".format(aktuelle_zeile)).value = information.beanstandung
    ws.range("G{}".format(aktuelle_zeile)).value = information.fahrzeugdaten.verkaufstyp
    ws.range("H{}".format(aktuelle_zeile)).value = information.fahrzeugdaten.laufleistung
    ws.range("I{}".format(aktuelle_zeile)).value = information.fahrzeugdaten.motor
    ws.range("J{}".format(aktuelle_zeile)).value = information.fahrzeugdaten.getriebe
    ws.range("K{}".format(aktuelle_zeile)).value = information.kundencodierung
    ws.range("L{}".format(aktuelle_zeile)).value = information.wekstattcodierung
    ws.range("M{}".format(aktuelle_zeile)).value = information.vz_nummer
    ws.range("N{}".format(aktuelle_zeile)).value = information.bnr
    ws.range("O{}".format(aktuelle_zeile)).value = information.ansprechpartner.name
    ws.range("P{}".format(aktuelle_zeile)).value = information.ansprechpartner.tel
    ws.range("Q{}".format(aktuelle_zeile)).value = information.postfach_filter
    ws.range("R{}".format(aktuelle_zeile)).value = information.url

    aktuelle_zeile += 1
    ws.autofit()
    # Speicher die Datei
    wb.save()
    # Schließe das Fenster
    app = xw.apps.active
    app.quit()
