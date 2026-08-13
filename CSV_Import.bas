Attribute VB_Name = "CSV_Import"
Option Explicit

' ================================================================
'  CSV IMPORT: MF_DACH_MAT.csv -> Artikel-Sheet
'  Vorhandene Artikel (gleiche ARTIKELNR) werden aktualisiert,
'  neue Artikel werden unten angefuegt.
'
'  VK-PREIS = Netto-EK x 1,35 (35% Aufschlag)
'
'  EINBINDEN: VBA-Editor oeffnen (Alt+F11) -> Datei ->
'             Datei importieren -> diese .bas-Datei auswaehlen
'  STARTEN:   Makro "CSV_Artikel_Importieren" ausfuehren
' ================================================================

Const VK_AUFSCHLAG As Double = 1.35   ' 35% Aufschlag auf Netto-EK

Sub CSV_Artikel_Importieren()

    Dim wsA As Worksheet
    Set wsA = LagerMakros.GetSheet("Artikel")
    If wsA Is Nothing Then
        MsgBox "Artikel-Blatt nicht gefunden!", vbCritical
        Exit Sub
    End If

    ' CSV-Pfad: gleicher Ordner wie diese Mappe
    Dim csvPfad As String
    csvPfad = ThisWorkbook.Path & "\MF_DACH_MAT.csv"

    ' Wenn nicht gefunden -> Dateidialog
    If Dir(csvPfad) = "" Then
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        fd.Title = "CSV-Datei auswaehlen (MF_DACH_MAT.csv)"
        fd.Filters.Add "CSV-Dateien", "*.csv"
        fd.AllowMultiSelect = False
        If fd.Show = True Then
            csvPfad = fd.SelectedItems(1)
        Else
            MsgBox "Abgebrochen.", vbInformation
            Exit Sub
        End If
    End If

    ' ---- Spalten im Artikel-Sheet dynamisch ermitteln ----
    Dim cNr    As Long: cNr    = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim cArt   As Long: cArt   = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim cTextB As Long: cTextB = LagerMakros.Spalte_Finden(wsA, "TextB")
    Dim cEAN   As Long: cEAN   = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim cEinh  As Long: cEinh  = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim cWG    As Long: cWG    = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim cEK    As Long: cEK    = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim cVK    As Long: cVK    = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")

    ' Pflicht: ARTIKELNR und ARTIKEL muessen vorhanden sein
    If cNr = 0 Or cArt = 0 Then
        MsgBox "Spalten ARTIKELNR oder ARTIKEL nicht gefunden!" & Chr(10) & _
               "Bitte sicherstellen, dass das Artikel-Sheet korrekt eingerichtet ist.", vbCritical
        Exit Sub
    End If

    ' ---- Letzte Datenzeile bestimmen ----
    Dim lastRow As Long
    lastRow = wsA.Cells(wsA.Rows.Count, cArt).End(xlUp).Row

    ' ---- Alle vorhandenen ARTIKELNR in Dictionary (fuer Duplikatcheck) ----
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lastRow
        Dim existNr As String
        existNr = Trim(CStr(wsA.Cells(i, cNr).Value))
        If existNr <> "" Then
            dict(existNr) = i
        End If
    Next i

    ' ---- CSV einlesen ----
    Dim ff As Integer: ff = FreeFile
    Open csvPfad For Input As #ff

    Dim csvZeile  As String
    Dim felder()  As String
    Dim istHeader As Boolean: istHeader = True
    Dim zNeu      As Long:    zNeu = 0
    Dim zAkt      As Long:    zAkt = 0
    Dim zFehler   As Long:    zFehler = 0

    ' CSV-Spaltenindizes (werden aus Headerzeile ermittelt, 0-basiert)
    Dim iNr     As Integer: iNr     = -1
    Dim iKurz1  As Integer: iKurz1  = -1
    Dim iKurz2  As Integer: iKurz2  = -1
    Dim iME     As Integer: iME     = -1
    Dim iEAN    As Integer: iEAN    = -1
    Dim iWG     As Integer: iWG     = -1
    Dim iNetto  As Integer: iNetto  = -1
    ' iEVP wird nicht mehr benoetigt (VK = EK x 1,35)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Do While Not EOF(ff)
        Line Input #ff, csvZeile
        felder = Split(csvZeile, ";")

        If istHeader Then
            ' Spaltenindizes aus der Kopfzeile ableiten
            Dim k As Integer
            For k = 0 To UBound(felder)
                Dim feldName As String
                feldName = Trim(felder(k))
                Select Case feldName
                    Case "Artikelnummer":  iNr    = k
                    Case "Kurztext1":      iKurz1 = k
                    Case "Kurztext2":      iKurz2 = k
                    Case "Mengeneinheit":  iME    = k
                    Case "EANNummer":      iEAN   = k
                    Case "Warengruppe":    iWG    = k
                    Case "Netto":          iNetto = k
                End Select
            Next k
            ' Pruefung: Pflichtfelder in CSV vorhanden?
            If iNr = -1 Or iKurz1 = -1 Then
                Close #ff
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                MsgBox "CSV-Format unbekannt: Spalten 'Artikelnummer' oder 'Kurztext1' fehlen!", vbCritical
                Exit Sub
            End If
            istHeader = False

        Else
            ' Datensatz verarbeiten
            If UBound(felder) < iNr Then GoTo WeiterCSV

            Dim artNr As String
            artNr = Trim(felder(iNr))
            If artNr = "" Then GoTo WeiterCSV

            ' Zielzeile bestimmen (neu oder update)
            Dim zRow As Long
            If dict.Exists(artNr) Then
                zRow = dict(artNr)
                zAkt = zAkt + 1
            Else
                lastRow = lastRow + 1
                zRow = lastRow
                dict(artNr) = zRow
                zNeu = zNeu + 1
            End If

            ' Daten in Artikel-Sheet schreiben
            wsA.Cells(zRow, cNr).Value = artNr

            If cArt > 0 And iKurz1 <= UBound(felder) Then
                wsA.Cells(zRow, cArt).Value = Trim(felder(iKurz1))
            End If

            If cTextB > 0 And iKurz2 >= 0 And iKurz2 <= UBound(felder) Then
                wsA.Cells(zRow, cTextB).Value = Trim(felder(iKurz2))
            End If

            If cEAN > 0 And iEAN >= 0 And iEAN <= UBound(felder) Then
                wsA.Cells(zRow, cEAN).Value = Trim(felder(iEAN))
            End If

            If cEinh > 0 And iME >= 0 And iME <= UBound(felder) Then
                wsA.Cells(zRow, cEinh).Value = Trim(felder(iME))
            End If

            If cWG > 0 And iWG >= 0 And iWG <= UBound(felder) Then
                wsA.Cells(zRow, cWG).Value = Trim(felder(iWG))
            End If

            If iNetto >= 0 And iNetto <= UBound(felder) Then
                Dim ekPreis As Double
                ekPreis = ZahlAusCSV(felder(iNetto))
                If cEK > 0 Then
                    wsA.Cells(zRow, cEK).Value = ekPreis
                End If
                ' VK-Preis = EK + 35% Aufschlag (auf 2 Nachkommastellen gerundet)
                If cVK > 0 Then
                    wsA.Cells(zRow, cVK).Value = Round(ekPreis * VK_AUFSCHLAG, 2)
                End If
            End If

        End If

WeiterCSV:
    Loop

    Close #ff

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Import abgeschlossen!" & Chr(10) & Chr(10) & _
           "Neue Artikel:       " & zNeu & Chr(10) & _
           "Aktualisiert:       " & zAkt & Chr(10) & _
           "Gesamt:             " & (zNeu + zAkt), _
           vbInformation, "CSV Import"

End Sub

' Hilfsfunktion: CSV-Zahlformat (Komma als Dezimaltrennzeichen) -> Double
Private Function ZahlAusCSV(s As String) As Double
    Dim cleaned As String
    cleaned = Replace(Trim(s), ",", ".")
    If IsNumeric(cleaned) Then
        ZahlAusCSV = CDbl(cleaned)
    Else
        ZahlAusCSV = 0
    End If
End Function
