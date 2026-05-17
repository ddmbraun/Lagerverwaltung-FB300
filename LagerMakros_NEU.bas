Attribute VB_Name = "LagerMakros"
Option Explicit

' ================================================================
'  EINSTELLUNGEN  -  hier anpassen
' ================================================================
Const ZEBRA_DRUCKER   As String = "ZDesigner GK420d"   ' Genauer Druckername in Windows
Const BENUTZER        As String = "Frank"
Const INV_DATEN_START As Long   = 6                    ' Inventurliste: Datenzeilen ab Zeile
Const BACKUP_MAX      As Long   = 4                    ' Maximale Anzahl Sicherungskopien
Const ART_HEADER_ROW  As Long   = 4                    ' Artikel-Sheet: Zeile mit Spaltenköpfen
Const ART_DATA_START  As Long   = 5                    ' Artikel-Sheet: erste Datenzeile

' InvSuche-Spalten - 11-Spalten-Struktur (A leer, B-K Daten)
Const INV_COL_NR      As Long = 2   ' B: Laufnummer
Const INV_COL_ARTNR   As Long = 3   ' C: Art.-Nr.
Const INV_COL_ARTIKEL As Long = 4   ' D: Artikelbezeichnung
Const INV_COL_SOLL    As Long = 5   ' E: SOLL-Bestand
Const INV_COL_IST     As Long = 6   ' F: IST / GEZAEHLT (Benutzereingabe)
Const INV_COL_DIFF    As Long = 7   ' G: DIFFERENZ (auto)
Const INV_COL_VK      As Long = 8   ' H: VK-Preis  (Button SUCHEN in Zeile 3)
Const INV_COL_EK      As Long = 9   ' I: EK-Preis  (Button LEEREN in Zeile 3)
Const INV_COL_EAN     As Long = 10  ' J: EAN13
Const INV_COL_LAGER   As Long = 11  ' K: Lagerort  (Button UEBERNEHMEN in Zeile 3)
Const INV_COL_ATTR    As Long = 12  ' L: Attribut

' Schnellansicht-Spalten (A leer, B-K Daten)
Const SA_COL_NR    As Long = 2   ' B: Nr.
Const SA_COL_ARTNR As Long = 3   ' C: Art.-Nr.
Const SA_COL_ART   As Long = 4   ' D: Artikelbezeichnung
Const SA_COL_EAN   As Long = 5   ' E: EAN13
Const SA_COL_VK    As Long = 6   ' F: VK-Preis
Const SA_COL_BEST  As Long = 7   ' G: Bestand (editierbar)
Const SA_COL_EINH  As Long = 8   ' H: Einheit
Const SA_COL_LAG   As Long = 9   ' I: Lagerort
Const SA_COL_WG    As Long = 10  ' J: Warengruppe
Const SA_COL_ATTR  As Long = 11  ' K: Attribut

' Globale Variablen
Public g_LetzteZeile           As Long   ' Zuletzt angeklickte Zeile (Artikel-Sheet)
Public g_ReturnSheet            As String ' Wohin ArtikelDetail zurückspringt ("Schnell" oder "Artikel")

' ================================================================
'  NOTFALL-RESET  (bei Absturz / Einfrieren ausführen via Alt+F8)
' ================================================================
Sub Reset_Alles()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    MsgBox "Reset OK – alles wieder aktiv.", vbInformation, "Reset"
End Sub

' ================================================================
'  SHEET SUCHEN  (robust, ohne Umlaut-Probleme)
' ================================================================
Function GetSheet(suchbegriff As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, suchbegriff, vbTextCompare) > 0 Then
            Set GetSheet = ws : Exit Function
        End If
    Next ws
    Set GetSheet = Nothing
End Function

Function SheetListe() As String
    Dim ws As Worksheet, s As String
    For Each ws In ThisWorkbook.Sheets
        s = s & ws.Name & Chr(10)
    Next ws
    SheetListe = s
End Function

' ================================================================
'  SPALTE FINDEN  (sucht in Headerzeile)
' ================================================================
Function Spalte_Finden(ws As Worksheet, headerName As String) As Long
    Dim hRow As Long
    hRow = IIf(InStr(1, ws.Name, "rtikel", vbTextCompare) > 0, ART_HEADER_ROW, 1)
    Dim lastCol As Long
    lastCol = ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column
    Dim i As Long
    For i = 1 To lastCol
        If InStr(1, CStr(ws.Cells(hRow, i).Value), headerName, vbTextCompare) > 0 Then
            Spalte_Finden = i : Exit Function
        End If
    Next i
    Spalte_Finden = 0
End Function

' ================================================================
'  WERKZEUGZEILEN-HANDLER  (Artikel-Sheet)
'  Zeile 1 = Suchfeld-Zeile (Suche + Leeren-Button)
'  Zeile 2 = Button-Zeile   (Zu/Abgang, Etikett, EK, Filter, Schnell, Neu)
'  Zeile >= ART_DATA_START  = Datenzeile (gelb markieren)
' ================================================================
Sub Toolbar_Handler(ByVal Target As Range)
    Dim ws As Worksheet : Set ws = Target.Worksheet

    ' Datenzeile angeklickt → nur vorherige gelbe Zeile löschen, neue markieren
    If Target.Row >= ART_DATA_START Then
        ' Leere Zeilen NICHT gelb markieren
        If Trim(CStr(ws.Cells(Target.Row, 2).Value)) = "" Then Exit Sub
        ' Gleiche Zeile nochmal → nichts tun
        If Target.Row = g_LetzteZeile Then Exit Sub

        Application.EnableEvents = False
        Application.ScreenUpdating = False
        ' Nur die zuletzt markierte Zeile zurücksetzen (kein Scan!)
        If g_LetzteZeile >= ART_DATA_START Then
            ws.Rows(g_LetzteZeile).Interior.ColorIndex = xlNone
        End If
        g_LetzteZeile = Target.Row
        ws.Rows(g_LetzteZeile).Interior.Color = RGB(255, 255, 153)
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Zeile 1: Titelzeile → nichts tun
    If Target.Row = 1 Then Exit Sub

    ' Zeile 2: Suchfeld-Zeile → Suchen / Leeren / Cursor im Suchfeld
    If Target.Row = 2 Then
        Dim btn1 As String : btn1 = UCase(Trim(CStr(Target.Cells(1, 1).Value)))
        Select Case btn1
            Case "SUCHEN"  : Artikelblatt_Suche
            Case "LEEREN"  : Artikelblatt_FilterLoeschen
            Case Else
                ' Klick ins Suchfeld → nichts tun (SendKeys entfernt, war Absturzursache)
        End Select
        Exit Sub
    End If

    ' Zeile 3: Button-Zeile → Funktionsbuttons (text-basiert, robust)
    If Target.Row = 3 Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Dim btn2 As String : btn2 = UCase(Trim(CStr(Target.Cells(1, 1).Value)))
        Select Case btn2
            Case "ZU-/ABGANG", "ZU/ABGANG", "ZUABGANG", "BUCHEN"    : ZuAbgang_Buchen
            Case "ETIKETT", "ETIKETT DRUCKEN"                        : Etikett_Drucken
            Case "EK", "EK-PREIS", "EK EINBLENDEN", "EK AUSBLENDEN", "EK EINBL.", "EK AUSBL." : EK_Toggle
            Case "FILTER LÖSCHEN", "FILTER LOESCHEN", "FILTER"      : Filter_Loeschen
            Case "SCHNELLANSICHT", "SCHNELL"                         : GetSheet("Schnell").Activate
            Case "NEUER ARTIKEL", "NEU"                              : NeuerArtikel
        End Select
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

' ================================================================
'  ZU-/ABGANG BUCHEN
' ================================================================
Sub ZuAbgang_Buchen()
    If g_LetzteZeile < ART_DATA_START Then
        MsgBox "Bitte zuerst eine Artikelzeile anklicken.", vbInformation, "Hinweis"
        Exit Sub
    End If

    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim wsZ As Worksheet : Set wsZ = GetSheet("Abg")
    Dim wsS As Worksheet : Set wsS = GetSheet("Schnell")

    If wsA Is Nothing Or wsZ Is Nothing Then
        MsgBox "Pflicht-Sheets nicht gefunden!" & Chr(10) & SheetListe(), vbCritical
        Exit Sub
    End If

    Dim zeile As Long : zeile = g_LetzteZeile

    Dim cEAN   As Long : cEAN   = Spalte_Finden(wsA, "EAN13")
    Dim cArt   As Long : cArt   = Spalte_Finden(wsA, "ARTIKEL")
    Dim cAnz   As Long : cAnz   = Spalte_Finden(wsA, "ANZAHL")
    Dim cNr    As Long : cNr    = Spalte_Finden(wsA, "ARTIKELNR")
    Dim cLager As Long : cLager = Spalte_Finden(wsA, "LAGERORT")
    Dim cVK    As Long : cVK    = Spalte_Finden(wsA, "VK-PREIS")
    Dim cWA    As Long : cWA    = Spalte_Finden(wsA, "WAStck")
    Dim cWE    As Long : cWE    = Spalte_Finden(wsA, "WEstck")

    Dim ean     As String  : ean     = Trim(CStr(wsA.Cells(zeile, cEAN).Value))
    Dim artikel As String  : artikel = Trim(wsA.Cells(zeile, cArt).Value)
    Dim artNr   As String  : artNr   = Trim(wsA.Cells(zeile, cNr).Value)
    Dim lager   As String  : lager   = Trim(wsA.Cells(zeile, cLager).Value)
    Dim aktuell As Double  : aktuell = Val(wsA.Cells(zeile, cAnz).Value)

    Dim eingabe As String
    eingabe = InputBox( _
        "Artikel:  " & artikel & Chr(10) & _
        "ArtNr:    " & artNr & Chr(10) & _
        "Aktueller Bestand: " & Format(aktuell, "0") & " Stk" & Chr(10) & Chr(10) & _
        "Menge eingeben:" & Chr(10) & _
        "  Zugang  →  positive Zahl  (z.B.  5)" & Chr(10) & _
        "  Abgang  →  negative Zahl  (z.B.  -3)", _
        "Zu-/Abgang buchen")

    If eingabe = "" Then Exit Sub
    Dim menge As Double : menge = Val(eingabe)
    If menge = 0 Then MsgBox "Ungültige Eingabe.", vbExclamation : Exit Sub

    Dim typ As String   : typ           = IIf(menge > 0, "Zugang", "Abgang")
    Dim neu As Double   : neu           = aktuell + menge

    ' 1) Artikel-Sheet: Bestand + WA/WE-Stck + Datum aktualisieren
    wsA.Cells(zeile, cAnz).Value = neu
    If menge > 0 Then
        If cWE > 0 Then wsA.Cells(zeile, cWE).Value = Val(wsA.Cells(zeile, cWE).Value) + menge
        ' WE-Datum setzen
        Dim cWEDat As Long : cWEDat = Spalte_Finden(wsA, "WE-Datum")
        If cWEDat > 0 Then
            wsA.Cells(zeile, cWEDat).Value = Now()
            wsA.Cells(zeile, cWEDat).NumberFormat = "DD.MM.YYYY HH:MM"
        End If
    Else
        If cWA > 0 Then wsA.Cells(zeile, cWA).Value = Val(wsA.Cells(zeile, cWA).Value) + Abs(menge)
    End If
    Dim cWADat As Long : cWADat = Spalte_Finden(wsA, "WA-Datum")
    If cWADat > 0 And menge < 0 Then
        wsA.Cells(zeile, cWADat).Value = Now()
        wsA.Cells(zeile, cWADat).NumberFormat = "DD.MM.YYYY HH:MM"
    End If

    ' 2) Schnellansicht synchronisieren
    Schnellansicht_BestandSync ean, neu

    ' 3) Bewegungsblatt eintragen
    Dim nRow As Long : nRow = wsZ.Cells(wsZ.Rows.Count, 1).End(xlUp).Row + 1
    With wsZ
        .Cells(nRow, 1).Value = Now()          : .Cells(nRow, 1).NumberFormat = "DD.MM.YYYY HH:MM"
        .Cells(nRow, 2).Value = ean
        .Cells(nRow, 3).Value = artNr
        .Cells(nRow, 4).Value = artikel
        .Cells(nRow, 5).Value = Abs(menge)
        .Cells(nRow, 6).Value = typ
        .Cells(nRow, 7).Value = lager
        .Cells(nRow, 8).Value = BENUTZER
    End With

    MsgBox typ & ": " & Format(Abs(menge), "0") & " Stk" & Chr(10) & _
           "Neuer Bestand: " & Format(neu, "0") & " Stk", vbInformation, "Buchung OK"
End Sub

' ================================================================
'  ETIKETT DRUCKEN  (Zebra GK420d, ZPL-Direktdruck)
' ================================================================
Sub Etikett_Drucken()
    If g_LetzteZeile < ART_DATA_START Then
        MsgBox "Bitte zuerst eine Artikelzeile anklicken.", vbInformation, "Hinweis"
        Exit Sub
    End If

    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim zeile As Long : zeile = g_LetzteZeile

    Dim cEAN   As Long : cEAN   = Spalte_Finden(wsA, "EAN13")
    Dim cArt   As Long : cArt   = Spalte_Finden(wsA, "ARTIKEL")
    Dim cVK    As Long : cVK    = Spalte_Finden(wsA, "VK-PREIS")
    Dim cNr    As Long : cNr    = Spalte_Finden(wsA, "ARTIKELNR")
    Dim cTextB As Long : cTextB = Spalte_Finden(wsA, "TextB")

    Dim ean     As String : ean     = Trim(CStr(wsA.Cells(zeile, cEAN).Value))
    Dim artikel As String : artikel = Trim(wsA.Cells(zeile, cArt).Value)
    Dim vkPreis As String : vkPreis = Format(wsA.Cells(zeile, cVK).Value, "0.00") & " EUR"
    Dim artNr   As String : artNr   = Trim(wsA.Cells(zeile, cNr).Value)
    Dim textB   As String
    If cTextB > 0 Then textB = Trim(wsA.Cells(zeile, cTextB).Value)

    If ean = "" Then
        MsgBox "Kein EAN vorhanden.", vbExclamation : Exit Sub
    End If

    ' Etikettengrösse wählen
    Dim groesse As Integer
    groesse = MsgBox("Etikettengrösse wählen:" & Chr(10) & Chr(10) & _
                     "[ Ja ]  =  70 x 38 mm  (Standard-Regal-Etikett)" & Chr(10) & _
                     "[ Nein ] =  30 x 50 mm  (Klein-Etikett)", _
                     vbYesNoCancel + vbQuestion, "Etikettengrösse")
    If groesse = vbCancel Then Exit Sub

    Dim zpl As String

    If groesse = vbYes Then
        ' --- 70 x 38 mm (560 x 304 Dots bei 203dpi) ---
        Dim zeile1 As String : zeile1 = Left(artikel, 35)
        Dim zeile2 As String : zeile2 = ""
        If Len(artikel) > 35 Then
            zeile2 = Mid(artikel, 36, 35)
        ElseIf textB <> "" Then
            zeile2 = Left(textB, 35)
        End If

        zpl = "^XA" & Chr(10)
        zpl = zpl & "^MMT^PW560^LL304^LS0" & Chr(10)
        zpl = zpl & "^FT12,40^A0N,35,35^FH\^FD" & zeile1 & "^FS" & Chr(10)
        If zeile2 <> "" Then
            zpl = zpl & "^FT12,80^A0N,32,32^FH\^FD" & zeile2 & "^FS" & Chr(10)
        End If
        zpl = zpl & "^FT12,118^A0N,28,28^FH\^FD" & artNr & "^FS" & Chr(10)
        zpl = zpl & "^FT12,152^A0N,34,34^FH\^FDBrutto " & vkPreis & "^FS" & Chr(10)
        zpl = zpl & "^FT35,295^BCN,80,Y,N,N^FD" & ean & "^FS" & Chr(10)
        zpl = zpl & "^PQ1^XZ"
    Else
        ' --- 30 x 50 mm (240 x 400 Dots bei 203dpi) ---
        Dim art30 As String : art30 = Left(artikel, 22)
        Dim art30b As String : art30b = ""
        If Len(artikel) > 22 Then art30b = Mid(artikel, 23, 22)

        zpl = "^XA" & Chr(10)
        zpl = zpl & "^MMT^PW240^LL400^LS0" & Chr(10)
        zpl = zpl & "^FT5,32^A0N,28,28^FH\^FD" & art30 & "^FS" & Chr(10)
        If art30b <> "" Then
            zpl = zpl & "^FT5,62^A0N,26,26^FH\^FD" & art30b & "^FS" & Chr(10)
        End If
        zpl = zpl & "^FT5,95^A0N,24,24^FH\^FD" & artNr & "^FS" & Chr(10)
        zpl = zpl & "^FT5,125^A0N,30,30^FH\^FDBrutto " & vkPreis & "^FS" & Chr(10)
        zpl = zpl & "^FT15,380^BCN,90,Y,N,N^FD" & ean & "^FS" & Chr(10)
        zpl = zpl & "^PQ1^XZ"
    End If

    ' ZPL in Temp-Datei schreiben und an Drucker senden
    Dim tmpDatei As String : tmpDatei = Environ("TEMP") & "\zebra_lager.zpl"
    Dim ff As Integer : ff = FreeFile
    Open tmpDatei For Output As #ff
    Print #ff, zpl
    Close #ff

    Dim oShell As Object : Set oShell = CreateObject("WScript.Shell")
    Dim ret As Long
    ret = oShell.Run("cmd /c copy /b """ & tmpDatei & """ """ & ZEBRA_DRUCKER & """", 0, True)

    If ret <> 0 Then
        MsgBox "Druckfehler (Code " & ret & ")" & Chr(10) & Chr(10) & _
               "Druckername prüfen: '" & ZEBRA_DRUCKER & "'" & Chr(10) & _
               "Windows: Einstellungen → Drucker & Scanner" & Chr(10) & _
               "→ Dort den genauen Namen des Zebra-Druckers eintragen (Const ZEBRA_DRUCKER im VBA).", _
               vbExclamation, "Druckfehler"
    Else
        MsgBox "Etikett gesendet!" & Chr(10) & _
               "Drucker: " & ZEBRA_DRUCKER & Chr(10) & _
               "Artikel: " & artikel & Chr(10) & _
               "EAN: " & ean, vbInformation, "Etikett gedruckt"
    End If
End Sub

' ================================================================
'  NEUER ARTIKEL
' ================================================================
Sub NeuerArtikel()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim wsS As Worksheet : Set wsS = GetSheet("Schnell")
    If wsA Is Nothing Then Exit Sub

    ' Spalten per Header suchen (position-unabhängig, Spalte A bleibt leer)
    Dim cEAN  As Long : cEAN  = Spalte_Finden(wsA, "EAN13")
    Dim cArt  As Long : cArt  = Spalte_Finden(wsA, "ARTIKEL")
    Dim cNr   As Long : cNr   = Spalte_Finden(wsA, "ARTIKELNR")
    Dim cVK   As Long : cVK   = Spalte_Finden(wsA, "VK-PREIS")
    Dim cAnz  As Long : cAnz  = Spalte_Finden(wsA, "ANZAHL")
    Dim cEinh As Long : cEinh = Spalte_Finden(wsA, "EINHEIT")
    Dim cWG   As Long : cWG   = Spalte_Finden(wsA, "WARENGRUPPE")
    Dim cMwSt As Long : cMwSt = Spalte_Finden(wsA, "MwSt")

    If cEAN = 0 Or cArt = 0 Then
        MsgBox "Spaltenköpfe nicht gefunden (EAN13 / ARTIKEL)." & Chr(10) & _
               "Bitte Artikelblatt_FelderErgaenzen ausführen.", vbExclamation
        Exit Sub
    End If

    Dim ean    As String : ean    = InputBox("EAN13 (13-stellig):", "Neuer Artikel - Schritt 1/5")
    If ean = "" Then Exit Sub
    Dim artNam As String : artNam = InputBox("Artikelbezeichnung:", "Neuer Artikel - Schritt 2/5")
    If artNam = "" Then Exit Sub
    Dim artNr  As String : artNr  = InputBox("Artikelnummer:", "Neuer Artikel - Schritt 3/5")
    Dim vkStr  As String : vkStr  = InputBox("VK-Preis (€):", "Neuer Artikel - Schritt 4/5")
    Dim wgStr  As String : wgStr  = InputBox("Warengruppe:", "Neuer Artikel - Schritt 5/5")

    Dim vk As Double : vk = Val(Replace(vkStr, ",", "."))

    ' Nächste freie Zeile (ab Spalte B suchen, da A immer leer)
    Dim nRow As Long
    nRow = wsA.Cells(wsA.Rows.Count, cArt).End(xlUp).Row + 1
    If nRow < ART_DATA_START Then nRow = ART_DATA_START

    ' Daten schreiben (Spalte A bleibt immer leer!)
    If cEAN  > 0 Then wsA.Cells(nRow, cEAN).Value  = ean
    If cArt  > 0 Then wsA.Cells(nRow, cArt).Value  = artNam
    If cNr   > 0 Then wsA.Cells(nRow, cNr).Value   = artNr
    If cVK   > 0 Then wsA.Cells(nRow, cVK).Value   = vk
    If cAnz  > 0 Then wsA.Cells(nRow, cAnz).Value  = 0
    If cEinh > 0 Then wsA.Cells(nRow, cEinh).Value = "Stk"
    If cWG   > 0 Then wsA.Cells(nRow, cWG).Value   = wgStr
    If cMwSt > 0 Then wsA.Cells(nRow, cMwSt).Value = 19

    ' Schnellansicht spiegeln
    If Not wsS Is Nothing Then
        Dim cSArt  As Long : cSArt  = Spalte_Finden(wsS, "ARTIKEL")
        Dim cSEAN  As Long : cSEAN  = Spalte_Finden(wsS, "EAN13")
        Dim cSVK   As Long : cSVK   = Spalte_Finden(wsS, "VK-PREIS")
        Dim cSAnz  As Long : cSAnz  = Spalte_Finden(wsS, "ANZAHL")
        Dim cSEinh As Long : cSEinh = Spalte_Finden(wsS, "EINHEIT")
        If cSArt > 0 Then
            Dim sRow As Long : sRow = wsS.Cells(wsS.Rows.Count, cSArt).End(xlUp).Row + 1
            If cSEAN  > 0 Then wsS.Cells(sRow, cSEAN).Value  = ean
            If cSArt  > 0 Then wsS.Cells(sRow, cSArt).Value  = artNam
            If cSVK   > 0 Then wsS.Cells(sRow, cSVK).Value   = vk
            If cSAnz  > 0 Then wsS.Cells(sRow, cSAnz).Value  = 0
            If cSEinh > 0 Then wsS.Cells(sRow, cSEinh).Value = "Stk"
        End If
    End If

    MsgBox "Artikel angelegt: " & artNam & Chr(10) & _
           "EAN: " & ean & Chr(10) & _
           "Zeile: " & nRow, vbInformation, "Neuer Artikel"
End Sub

' ================================================================
'  EK-PREIS EIN-/AUSBLENDEN
' ================================================================
Sub EK_Toggle()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim ekSpalte As Long : ekSpalte = Spalte_Finden(wsA, "EK-PREIS")
    If ekSpalte = 0 Then Exit Sub
    wsA.Columns(ekSpalte).Hidden = Not wsA.Columns(ekSpalte).Hidden
    ' EK-Button (Zeile 3, Spalte 7) Text aktualisieren
    With wsA.Cells(3, 7)
        .Value = IIf(wsA.Columns(ekSpalte).Hidden, "EK einbl.", "EK ausbl.")
    End With
End Sub

' ================================================================
'  FILTER LÖSCHEN  (Artikel-Sheet)
' ================================================================
Sub Filter_Loeschen()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    If g_LetzteZeile >= ART_DATA_START Then wsA.Rows(g_LetzteZeile).Interior.ColorIndex = xlNone
    If wsA.AutoFilterMode Then wsA.AutoFilter.ShowAllData
    wsA.Cells(ART_DATA_START, 1).Select
    g_LetzteZeile = 0
End Sub

' ================================================================
'  SCHNELLANSICHT - HANDLER  (Zeile 2 = Suchzeile)
' ================================================================
Sub Schnellansicht_Handler(ByVal Target As Range)
    If Target.Row <> 2 Then Exit Sub
    On Error Resume Next
    Dim btn As String : btn = UCase(Trim(CStr(Target.Cells(1, 1).Value)))
    On Error GoTo 0
    Select Case btn
        Case "SUCHEN"       : Schnellansicht_Suchen
        Case "LEEREN"       : Schnellansicht_FilterLoeschen
        Case "AKTUALISIEREN": Schnellansicht_Aktualisieren
    End Select
End Sub

' ================================================================
'  SCHNELLANSICHT - SUCHE MIT MEHRWORT-POPUP
' ================================================================
Sub Schnellansicht_Suchen()
    Dim wsS As Worksheet : Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub
    Dim such As String : such = Trim(wsS.Cells(2, 3).Value)  ' Suchfeld C2
    If such = "" Then Schnellansicht_FilterLoeschen : Exit Sub
    Dim woerter()  As String  : woerter   = Split(LCase(such), " ")
    Dim nurZahlen  As Boolean : nurZahlen = (such = CStr(Val(such)) And Val(such) > 0)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim lastRow As Long : lastRow = wsS.Cells(wsS.Rows.Count, SA_COL_ART).End(xlUp).Row
    Dim treffer As Long : treffer = 0
    Dim i As Long, w As Integer, passt As Boolean
    For i = 4 To lastRow
        Dim suchIn As String
        If nurZahlen Then
            suchIn = LCase(CStr(wsS.Cells(i, SA_COL_ART).Value) & " " & _
                          CStr(wsS.Cells(i, SA_COL_ARTNR).Value) & " " & _
                          CStr(wsS.Cells(i, SA_COL_EAN).Value))
        Else
            suchIn = LCase(CStr(wsS.Cells(i, SA_COL_ART).Value))
        End If
        passt = True
        For w = 0 To UBound(woerter)
            If Trim(woerter(w)) <> "" Then
                If InStr(suchIn, Trim(woerter(w))) = 0 Then passt = False : Exit For
            End If
        Next w
        wsS.Rows(i).Hidden = Not passt
        If passt Then treffer = treffer + 1
    Next i
    wsS.Cells(2, 10).Value = treffer & " Treffer"   ' J2
    wsS.Cells(1, 1).Select
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  SCHNELLANSICHT - FILTER LÖSCHEN
' ================================================================
Sub Schnellansicht_FilterLoeschen()
    Dim wsS As Worksheet : Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Dim lastRow As Long : lastRow = wsS.Cells(wsS.Rows.Count, SA_COL_ART).End(xlUp).Row
    Dim i As Long
    For i = 4 To lastRow
        wsS.Rows(i).Hidden = False
    Next i
    wsS.Cells(2, 3).Value = ""   ' Suchfeld C2
    wsS.Cells(2, 10).Value = ""  ' Treffer J2
    Application.EnableEvents = True
End Sub

' ================================================================
'  SCHNELLANSICHT - AKTUALISIEREN (aus Artikel-Sheet neu laden)
' ================================================================
Sub Schnellansicht_Aktualisieren()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim wsS As Worksheet : Set wsS = GetSheet("Schnell")
    If wsA Is Nothing Or wsS Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim cEAN  As Long : cEAN  = Spalte_Finden(wsA, "EAN13")
    Dim cArt  As Long : cArt  = Spalte_Finden(wsA, "ARTIKEL")
    Dim cNr   As Long : cNr   = Spalte_Finden(wsA, "ARTIKELNR")
    Dim cVK   As Long : cVK   = Spalte_Finden(wsA, "VK-PREIS")
    Dim cAnz  As Long : cAnz  = Spalte_Finden(wsA, "ANZAHL")
    Dim cEinh As Long : cEinh = Spalte_Finden(wsA, "EINHEIT")
    Dim cLag  As Long : cLag  = Spalte_Finden(wsA, "LAGERORT")
    Dim cWG   As Long : cWG   = Spalte_Finden(wsA, "WARENGRUPPE")
    Dim cAttr As Long : cAttr = Spalte_Finden(wsA, "ATTRIBUT")

    ' Autofilter entfernen
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False
    ' Alte Daten komplett löschen - UsedRange damit auch alte Spaltenstruktur erwischt
    Dim lastSvRow As Long
    lastSvRow = wsS.UsedRange.Row + wsS.UsedRange.Rows.Count - 1
    If lastSvRow >= 4 Then
        wsS.Rows("4:" & lastSvRow).ClearContents
        wsS.Rows("4:" & lastSvRow).ClearFormats
        wsS.Rows("4:" & lastSvRow).Hidden = False
    End If

    ' Neu einlesen
    Dim i As Long, sRow As Long : sRow = 4
    For i = ART_DATA_START To wsA.Cells(wsA.Rows.Count, cArt).End(xlUp).Row
        If wsA.Cells(i, cEAN).Value <> "" And wsA.Cells(i, cArt).Value <> "" Then
            wsS.Cells(sRow, SA_COL_NR).Value    = sRow - 3
            wsS.Cells(sRow, SA_COL_ARTNR).NumberFormat = "@"
            wsS.Cells(sRow, SA_COL_ARTNR).Value = CStr(wsA.Cells(i, cNr).Value)
            wsS.Cells(sRow, SA_COL_ART).Value   = wsA.Cells(i, cArt).Value
            wsS.Cells(sRow, SA_COL_EAN).NumberFormat = "@"
            wsS.Cells(sRow, SA_COL_EAN).Value   = CStr(wsA.Cells(i, cEAN).Value)
            wsS.Cells(sRow, SA_COL_VK).Value    = wsA.Cells(i, cVK).Value
            wsS.Cells(sRow, SA_COL_BEST).Value  = wsA.Cells(i, cAnz).Value
            wsS.Cells(sRow, SA_COL_EINH).Value  = wsA.Cells(i, cEinh).Value
            If cLag  > 0 Then wsS.Cells(sRow, SA_COL_LAG).Value  = wsA.Cells(i, cLag).Value
            If cWG   > 0 Then wsS.Cells(sRow, SA_COL_WG).Value   = wsA.Cells(i, cWG).Value
            If cAttr > 0 Then wsS.Cells(sRow, SA_COL_ATTR).Value = wsA.Cells(i, cAttr).Value
            ' Bestand 0 oder niedrig → Zeile einfärben
            Dim bestand As Double : bestand = Val(wsA.Cells(i, cAnz).Value)
            If bestand = 0 Then
                wsS.Rows(sRow).Interior.Color = RGB(255, 199, 206)
            ElseIf bestand <= 5 Then
                wsS.Rows(sRow).Interior.Color = RGB(255, 235, 156)
            End If
            sRow = sRow + 1
        End If
    Next i

    ' Header sicherstellen
    Schnellansicht_Headers wsS
    wsS.Cells(1, 1).Select
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Schnellansicht aktualisiert: " & sRow - 4 & " Artikel.", vbInformation, "Aktualisiert"
End Sub

' ================================================================
'  INVENTURLISTE - BEFUELLEN (aus Artikel-Sheet)
' ================================================================
Sub Inventurliste_Befuellen()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet : Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsI Is Nothing Then
        MsgBox "Sheet nicht gefunden!" & Chr(10) & SheetListe(), vbCritical : Exit Sub
    End If

    If MsgBox("Inventurliste jetzt mit allen aktiven Artikeln befüllen?" & Chr(10) & _
              "Bestehende Einträge werden überschrieben.", _
              vbYesNo + vbQuestion, "Inventurliste befüllen") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Daten ab Zeile INV_DATEN_START löschen
    Dim lastI As Long : lastI = wsI.Cells(wsI.Rows.Count, 2).End(xlUp).Row
    If lastI >= INV_DATEN_START Then
        wsI.Rows(INV_DATEN_START & ":" & lastI).ClearContents
    End If

    Dim cEAN  As Long : cEAN  = Spalte_Finden(wsA, "EAN13")
    Dim cArt  As Long : cArt  = Spalte_Finden(wsA, "ARTIKEL")
    Dim cLag  As Long : cLag  = Spalte_Finden(wsA, "LAGERORT")
    Dim cEK   As Long : cEK   = Spalte_Finden(wsA, "EK-PREIS")
    Dim cAnz  As Long : cAnz  = Spalte_Finden(wsA, "ANZAHL")

    Dim i As Long, iRow As Long : iRow = INV_DATEN_START
    For i = ART_DATA_START To wsA.Cells(wsA.Rows.Count, cArt).End(xlUp).Row
        If wsA.Cells(i, cEAN).Value <> "" And wsA.Cells(i, cArt).Value <> "" Then
            wsI.Cells(iRow, 1).Value = iRow - INV_DATEN_START + 1   ' Nr
            wsI.Cells(iRow, 2).Value = wsA.Cells(i, cEAN).Value     ' EAN
            wsI.Cells(iRow, 3).Value = wsA.Cells(i, cArt).Value     ' Artikel
            wsI.Cells(iRow, 4).Value = wsA.Cells(i, cLag).Value     ' Lagerort
            wsI.Cells(iRow, 5).Value = wsA.Cells(i, cEK).Value      ' EK-Preis
            wsI.Cells(iRow, 6).Value = wsA.Cells(i, cAnz).Value     ' SOLL
            wsI.Cells(iRow, 7).Value = ""                           ' GEZÄHLT (leer)
            wsI.Cells(iRow, 8).Formula = "=IF(G" & iRow & "="""","""",G" & iRow & "-F" & iRow & ")"
            wsI.Cells(iRow, 9).Formula = "=IF(G" & iRow & "="""","""",G" & iRow & "*E" & iRow & ")"
            iRow = iRow + 1
        End If
    Next i

    wsI.Cells(2, 2).Value = Date
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Inventurliste befüllt: " & iRow - INV_DATEN_START & " Artikel.", vbInformation, "Bereit"
End Sub

' ================================================================
'  INVENTURLISTE - ÜBERNEHMEN (Korrekturen → Bewegungen + Artikel)
' ================================================================
Sub Inventurliste_Uebernehmen()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet : Set wsI = GetSheet("Inventur")
    Dim wsZ As Worksheet : Set wsZ = GetSheet("Abg")
    If wsA Is Nothing Or wsI Is Nothing Or wsZ Is Nothing Then
        MsgBox "Sheet nicht gefunden!", vbCritical : Exit Sub
    End If

    If MsgBox("Inventur-Korrekturen jetzt übernehmen?" & Chr(10) & _
              "Alle Artikel mit Differenz <> 0 werden im Bewegungsblatt eingetragen" & Chr(10) & _
              "und der Bestand im Artikel-Sheet wird korrigiert.", _
              vbYesNo + vbQuestion, "Inventur übernehmen") = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    Dim cEAN As Long : cEAN = Spalte_Finden(wsA, "EAN13")
    Dim cAnz As Long : cAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim wsS  As Worksheet : Set wsS = GetSheet("Schnell")

    Dim lastI As Long : lastI = wsI.Cells(wsI.Rows.Count, 2).End(xlUp).Row
    Dim korrekturen As Long : korrekturen = 0

    Dim i As Long
    For i = INV_DATEN_START To lastI
        Dim ean      As String  : ean    = CStr(wsI.Cells(i, 2).Value)
        Dim soll     As Double  : soll   = Val(wsI.Cells(i, 6).Value)
        Dim gezaehlt As String  : gezaehlt = CStr(wsI.Cells(i, 7).Value)
        If ean = "" Or gezaehlt = "" Then GoTo WeiterI

        Dim istMenge As Double : istMenge = Val(gezaehlt)
        Dim diff     As Double : diff     = istMenge - soll
        If diff = 0 Then GoTo WeiterI

        ' Bewegung schreiben
        Dim nRow As Long : nRow = wsZ.Cells(wsZ.Rows.Count, 1).End(xlUp).Row + 1
        wsZ.Cells(nRow, 1).Value = Now()
        wsZ.Cells(nRow, 1).NumberFormat = "DD.MM.YYYY HH:MM"
        wsZ.Cells(nRow, 2).Value = ean
        wsZ.Cells(nRow, 3).Value = wsI.Cells(i, 3).Value
        wsZ.Cells(nRow, 4).Value = wsI.Cells(i, 3).Value
        wsZ.Cells(nRow, 5).Value = Abs(diff)
        wsZ.Cells(nRow, 6).Value = IIf(diff > 0, "Inventur-Zugang", "Inventur-Abgang")
        wsZ.Cells(nRow, 7).Value = wsI.Cells(i, 4).Value
        wsZ.Cells(nRow, 8).Value = BENUTZER
        wsZ.Cells(nRow, 9).Value = "Inventur " & Format(Date, "DD.MM.YYYY")

        ' Bestand in Artikel-Sheet korrigieren
        Dim j As Long
        For j = ART_DATA_START To wsA.Cells(wsA.Rows.Count, cEAN).End(xlUp).Row
            If CStr(wsA.Cells(j, cEAN).Value) = ean Then
                wsA.Cells(j, cAnz).Value = istMenge
                Schnellansicht_BestandSync ean, istMenge
                Exit For
            End If
        Next j
        korrekturen = korrekturen + 1
WeiterI:
    Next i

    Application.ScreenUpdating = True
    MsgBox "Inventur übernommen: " & korrekturen & " Korrekturen eingetragen.", _
           vbInformation, "Inventur abgeschlossen"
End Sub

' ================================================================
'  INVSУCHE - HANDLER
'  Buttons werden anhand ihres Textes erkannt (robust)
'  Zeile 4+  →  Artikel anklicken → Eingabeblatt öffnen
' ================================================================
Sub InvSuche_Handler(ByVal Target As Range)
    ' Buttons in Zeile 2 - Erkennung anhand Zelleninhalt
    If Target.Row <> 2 Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Cleanup
    Dim btn As String : btn = UCase(Trim(CStr(Target.Value)))
    Select Case btn
        Case "SUCHEN"                        : InvSuche_Suchen
        Case "LEEREN"                        : InvSuche_Leeren
        Case "UEBERNEHMEN", "ÜBERNEHMEN"    : InvSuche_Uebernehmen
        Case "INVENTUR STARTEN"              : Inventur_Starten
        Case "GEZAEHLT ANZEIGEN"             : InvDaten_GezaehltAnzeigen
    End Select
Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  INVSУCHE - DOPPELKLICK HANDLER (Artikel → Eingabeblatt)
'  Wird aus Worksheet_BeforeDoubleClick des InvSuche-Sheets
'  aufgerufen.
' ================================================================
Sub InvSuche_DoppelklickHandler(ByVal Target As Range, Cancel As Boolean)
    If Target.Row < 4 Then Exit Sub
    Cancel = True   ' Standard-Doppelklick (Editiermodus) unterdrücken
    If Target.Cells.Count <> 1 Then Exit Sub
    If Target.Worksheet.Cells(Target.Row, INV_COL_ARTIKEL).Value <> "" Then
        InvSuche_Artikelklick Target.Worksheet, Target.Row
    End If
End Sub

' ================================================================
'  INVSУCHE - GEZÄHLT-EINGABE VERARBEITEN  (Worksheet_Change)
'
'  Wird aufgerufen wenn Benutzer in Spalte 8 (GEZÄHLT) tippt.
'  → Differenz berechnen und Zeile einfärben.
' ================================================================
Sub InvSuche_GezaehltChange(ByVal Target As Range)
    ' Nur Spalte IST, Datenzeilen ab 4
    If Target.Column <> INV_COL_IST Then Exit Sub
    If Target.Row < 4 Then Exit Sub
    ' Mehrfachselektion ignorieren
    If Target.Cells.Count > 1 Then Exit Sub

    Dim ws As Worksheet : Set ws = Target.Worksheet
    Dim zeile As Long : zeile = Target.Row

    ' Artikel-Identifikation prüfen
    If ws.Cells(zeile, INV_COL_ARTIKEL).Value = "" Then Exit Sub

    Dim soll    As Double : soll    = Val(ws.Cells(zeile, INV_COL_SOLL).Value)
    Dim gezVal  As String : gezVal  = Trim(CStr(ws.Cells(zeile, INV_COL_IST).Value))

    Application.EnableEvents = False

    If gezVal = "" Then
        ' Noch nicht gezählt → Farbe zurücksetzen, Differenz löschen
        ws.Cells(zeile, INV_COL_DIFF).Value = ""
        ws.Rows(zeile).Interior.ColorIndex = xlNone
    Else
        Dim istW As Double : istW = Val(gezVal)
        Dim diff As Double : diff = istW - soll
        ' Differenz schreiben
        ws.Cells(zeile, INV_COL_DIFF).Value = diff
        ws.Cells(zeile, INV_COL_DIFF).NumberFormat = IIf(diff >= 0, "+0;-0;0", "0")
        ' Zeile einfärben
        Select Case True
            Case diff = 0   : ws.Rows(zeile).Interior.Color = RGB(198, 239, 206)  ' grün   = stimmt
            Case diff > 0   : ws.Rows(zeile).Interior.Color = RGB(255, 235, 156)  ' gelb   = mehr gezählt
            Case Else       : ws.Rows(zeile).Interior.Color = RGB(255, 199, 206)  ' rot    = weniger gezählt
        End Select
    End If

    Application.EnableEvents = True
End Sub

' ================================================================
'  INVSУCHE - SUCHEN
'  Schreibt Suchergebnisse ab Zeile 4 ins InvSuche-Sheet.
'  Bereits eingetragene GEZÄHLT-Werte (Spalte 8) werden
'  anhand der EAN wiederhergestellt.
' ================================================================
Sub InvSuche_Suchen()
    Dim wsS  As Worksheet : Set wsS  = GetSheet("InvSuch")
    Dim wsA  As Worksheet : Set wsA  = GetSheet("Artikel")
    If wsS Is Nothing Or wsA Is Nothing Then Exit Sub

    ' Suchbegriff lesen (Suchfeld = C2, Spalte 3)
    Dim such As String : such = Trim(wsS.Cells(2, 3).Value)
    If such = "" Then InvSuche_Leeren : Exit Sub

    Dim woerter()  As String  : woerter   = Split(LCase(such), " ")
    Dim nurZahlen  As Boolean : nurZahlen = (such = CStr(Val(such)) And Val(such) > 0)

    Dim cEAN  As Long : cEAN  = Spalte_Finden(wsA, "EAN13")
    Dim cArt  As Long : cArt  = Spalte_Finden(wsA, "ARTIKEL")
    Dim cAnz  As Long : cAnz  = Spalte_Finden(wsA, "ANZAHL")
    Dim cVK   As Long : cVK   = Spalte_Finden(wsA, "VK-PREIS")
    Dim cEK   As Long : cEK   = Spalte_Finden(wsA, "EK-PREIS")
    Dim cNr   As Long : cNr   = Spalte_Finden(wsA, "ARTIKELNR")
    Dim cLag  As Long : cLag  = Spalte_Finden(wsA, "LAGERORT")
    Dim cAttr As Long : cAttr = Spalte_Finden(wsA, "ATTRIBUT")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ── Alte Ergebnisse löschen (ab Zeile 4) ──
    Dim lastOld As Long : lastOld = wsS.Cells(wsS.Rows.Count, INV_COL_ARTIKEL).End(xlUp).Row
    If lastOld >= 4 Then
        wsS.Rows("4:" & lastOld).ClearContents
        wsS.Rows("4:" & lastOld).Interior.ColorIndex = xlNone
    End If

    ' ── Spaltenüberschriften in Zeile 3 sicherstellen ──
    InvSuche_Headers wsS

    ' ── Treffer schreiben ──
    Dim i As Long, w As Integer, passt As Boolean
    Dim sRow As Long : sRow = 4

    For i = ART_DATA_START To wsA.Cells(wsA.Rows.Count, cArt).End(xlUp).Row
        Dim suchIn As String
        If nurZahlen Then
            ' Zahlensuche → in Artikelname + Art.-Nr. + EAN
            suchIn = LCase(CStr(wsA.Cells(i, cArt).Value) & " " & _
                          CStr(wsA.Cells(i, cNr).Value) & " " & _
                          CStr(wsA.Cells(i, cEAN).Value))
        Else
            ' Textsuche → nur Artikelname (verhindert falsche Treffer über Art.-Nr.)
            suchIn = LCase(CStr(wsA.Cells(i, cArt).Value))
        End If
        passt = True
        For w = 0 To UBound(woerter)
            If Trim(woerter(w)) <> "" Then
                If InStr(suchIn, Trim(woerter(w))) = 0 Then passt = False : Exit For
            End If
        Next w

        If passt Then
            Dim eanVal As String : eanVal = Trim(CStr(wsA.Cells(i, cEAN).Value))
            ' B: Nr
            wsS.Cells(sRow, INV_COL_NR).Value = sRow - 3
            ' C: Art.-Nr. (CStr verhindert Sci-Notation)
            wsS.Cells(sRow, INV_COL_ARTNR).NumberFormat = "@"
            wsS.Cells(sRow, INV_COL_ARTNR).Value = CStr(wsA.Cells(i, cNr).Value)
            ' D: Artikel
            wsS.Cells(sRow, INV_COL_ARTIKEL).Value = wsA.Cells(i, cArt).Value
            ' E: SOLL
            wsS.Cells(sRow, INV_COL_SOLL).Value = wsA.Cells(i, cAnz).Value
            ' F: IST - aus InvDaten wiederherstellen (persistent über Sessions)
            Dim altGez As String : altGez = InvDaten_IstHolen(eanVal)
            wsS.Cells(sRow, INV_COL_IST).Value = altGez
            ' G: DIFFERENZ
            If altGez <> "" Then
                Dim altDiff As Double : altDiff = Val(altGez) - wsA.Cells(i, cAnz).Value
                wsS.Cells(sRow, INV_COL_DIFF).Value = altDiff
                wsS.Cells(sRow, INV_COL_DIFF).NumberFormat = IIf(altDiff >= 0, "+0;-0;0", "0")
                Select Case True
                    Case altDiff = 0 : wsS.Rows(sRow).Interior.Color = RGB(198, 239, 206)
                    Case altDiff > 0 : wsS.Rows(sRow).Interior.Color = RGB(255, 235, 156)
                    Case Else        : wsS.Rows(sRow).Interior.Color = RGB(255, 199, 206)
                End Select
            Else
                wsS.Cells(sRow, INV_COL_DIFF).Value = ""
            End If
            ' H: VK-Preis
            If cVK > 0 Then wsS.Cells(sRow, INV_COL_VK).Value = wsA.Cells(i, cVK).Value
            ' I: EK-Preis
            If cEK > 0 Then wsS.Cells(sRow, INV_COL_EK).Value = wsA.Cells(i, cEK).Value
            ' J: EAN (NumberFormat "@" verhindert Sci-Notation)
            wsS.Cells(sRow, INV_COL_EAN).NumberFormat = "@"
            wsS.Cells(sRow, INV_COL_EAN).Value = CStr(eanVal)
            ' K: Lagerort
            wsS.Cells(sRow, INV_COL_LAGER).Value = wsA.Cells(i, cLag).Value
            ' L: Attribut
            If cAttr > 0 Then wsS.Cells(sRow, INV_COL_ATTR).Value = wsA.Cells(i, cAttr).Value

            sRow = sRow + 1
        End If
    Next i

    wsS.Cells(2, 10).Value = (sRow - 4) & " Treffer"
    wsS.Cells(1, 1).Select   ' Ansicht zurück zu Anfang scrollen
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  INVSУCHE - SPALTENÜBERSCHRIFTEN (Zeile 3) setzen
' ================================================================
Private Sub InvSuche_Headers(wsS As Worksheet)
    ' Spalte A leer, Zeile 3 = reine Spaltenköpfe (keine Buttons)
    wsS.Cells(3, 1).Value = ""
    wsS.Cells(3, 1).Interior.ColorIndex = xlNone

    ' Spalten B(2)-K(11): alle dunkelgrau mit weisser Schrift
    Dim headers As Variant
    headers = Array("#", "Art.-Nr.", "Artikel", "SOLL", "IST / Gezaehlt", _
                    "DIFFERENZ", "VK-Preis", "EK-Preis", "EAN", "Lagerort", "Attribut")
    ' Index 0 = Spalte B(2)  …  Index 10 = Spalte L(12)
    Dim col As Long
    For col = 2 To 12
        With wsS.Cells(3, col)
            .Value = headers(col - 2)
            .Font.Bold = True
            .Font.Name = "Arial"
            .Font.Size = 12
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(64, 64, 64)
            .HorizontalAlignment = xlCenter
        End With
    Next col
End Sub

' ================================================================
'  INVSУCHE - LEEREN
' ================================================================
Sub InvSuche_Leeren()
    Dim wsS As Worksheet : Set wsS = GetSheet("InvSuch")
    If wsS Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Dim lastRow As Long : lastRow = wsS.Cells(wsS.Rows.Count, INV_COL_ARTIKEL).End(xlUp).Row
    If lastRow >= 4 Then
        wsS.Rows("4:" & lastRow).ClearContents
        wsS.Rows("4:" & lastRow).Interior.ColorIndex = xlNone
    End If
    wsS.Cells(2, 3).Value = ""
    wsS.Cells(2, 10).Value = ""
    Application.EnableEvents = True
End Sub

' ================================================================
'  INVSУCHE - ÜBERNEHMEN
'  Schreibt alle eingetragenen GEZÄHLT-Werte mit Differenz
'  in Bewegungsblatt und korrigiert den Bestand in Artikel-Sheet.
' ================================================================
Sub InvSuche_Uebernehmen()
    Dim wsS As Worksheet : Set wsS = GetSheet("InvSuch")
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim wsZ As Worksheet : Set wsZ = GetSheet("Abg")
    If wsS Is Nothing Or wsA Is Nothing Or wsZ Is Nothing Then
        MsgBox "Sheet nicht gefunden!" & Chr(10) & SheetListe(), vbCritical : Exit Sub
    End If

    ' Prüfen ob überhaupt Einträge vorhanden
    Dim lastRow As Long : lastRow = wsS.Cells(wsS.Rows.Count, INV_COL_ARTIKEL).End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "Keine Suchergebnisse vorhanden. Bitte zuerst suchen.", vbInformation : Exit Sub
    End If

    ' Zählen wie viele GEZÄHLT-Werte vorhanden
    Dim anzGez As Long : anzGez = 0
    Dim r As Long
    For r = 4 To lastRow
        If Trim(CStr(wsS.Cells(r, INV_COL_IST).Value)) <> "" Then anzGez = anzGez + 1
    Next r
    If anzGez = 0 Then
        MsgBox "Keine GEZÄHLT-Mengen eingetragen.", vbInformation : Exit Sub
    End If

    If MsgBox("Inventur-Korrekturen aus der Suche übernehmen?" & Chr(10) & _
              anzGez & " Artikel mit eingetragenen Mengen." & Chr(10) & Chr(10) & _
              "Abweichungen werden im Bewegungsblatt eingetragen" & Chr(10) & _
              "und der Bestand im Artikel-Sheet wird korrigiert.", _
              vbYesNo + vbQuestion, "InvSuche - Übernehmen") = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    Dim cEAN As Long : cEAN = Spalte_Finden(wsA, "EAN13")
    Dim cAnz As Long : cAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim wsSc As Worksheet : Set wsSc = GetSheet("Schnell")
    Dim korrekturen As Long : korrekturen = 0

    For r = 4 To lastRow
        Dim ean     As String : ean     = CStr(wsS.Cells(r, INV_COL_EAN).Value)
        Dim artikel As String : artikel = CStr(wsS.Cells(r, INV_COL_ARTIKEL).Value)
        Dim soll    As Double : soll    = Val(wsS.Cells(r, INV_COL_SOLL).Value)
        Dim gezStr  As String : gezStr  = Trim(CStr(wsS.Cells(r, INV_COL_IST).Value))
        Dim lager   As String : lager   = CStr(wsS.Cells(r, INV_COL_LAGER).Value)

        If artikel = "" Or gezStr = "" Then GoTo WeiterR

        Dim ist  As Double : ist  = Val(gezStr)
        Dim diff As Double : diff = ist - soll
        If diff = 0 Then GoTo WeiterR  ' Kein Unterschied → überspringen

        ' Bewegung schreiben
        Dim nRow As Long : nRow = wsZ.Cells(wsZ.Rows.Count, 1).End(xlUp).Row + 1
        wsZ.Cells(nRow, 1).Value = Now()
        wsZ.Cells(nRow, 1).NumberFormat = "DD.MM.YYYY HH:MM"
        wsZ.Cells(nRow, 2).Value = ean
        wsZ.Cells(nRow, 3).Value = wsS.Cells(r, INV_COL_ARTNR).Value
        wsZ.Cells(nRow, 4).Value = artikel
        wsZ.Cells(nRow, 5).Value = Abs(diff)
        wsZ.Cells(nRow, 6).Value = IIf(diff > 0, "Inventur-Zugang", "Inventur-Abgang")
        wsZ.Cells(nRow, 7).Value = lager
        wsZ.Cells(nRow, 8).Value = BENUTZER
        wsZ.Cells(nRow, 9).Value = "InvSuche " & Format(Date, "DD.MM.YYYY")

        ' Bestand im Artikel-Sheet aktualisieren
        Dim j As Long
        For j = ART_DATA_START To wsA.Cells(wsA.Rows.Count, cEAN).End(xlUp).Row
            If CStr(wsA.Cells(j, cEAN).Value) = ean Then
                wsA.Cells(j, cAnz).Value = ist
                Schnellansicht_BestandSync ean, ist
                Exit For
            End If
        Next j
        korrekturen = korrekturen + 1
WeiterR:
    Next r

    Application.ScreenUpdating = True
    MsgBox "Übernommen: " & korrekturen & " Korrekturen eingetragen." & Chr(10) & _
           "Bestände im Artikel-Sheet wurden aktualisiert.", _
           vbInformation, "InvSuche - Übernehmen abgeschlossen"
End Sub

' ================================================================
'  INVSУCHE - ARTIKEL ANKLICKEN → EINGABEBLATT ÖFFNEN
' ================================================================
Sub InvSuche_Artikelklick(ws As Worksheet, zeile As Long)
    If zeile < 4 Then Exit Sub
    Dim ean     As String : ean     = CStr(ws.Cells(zeile, INV_COL_EAN).Value)
    Dim artNr   As String : artNr   = ws.Cells(zeile, INV_COL_ARTNR).Value
    Dim artikel As String : artikel = ws.Cells(zeile, INV_COL_ARTIKEL).Value
    Dim soll    As Double : soll    = Val(ws.Cells(zeile, INV_COL_SOLL).Value)
    Dim vk      As Double : vk      = Val(ws.Cells(zeile, INV_COL_VK).Value)
    Dim lager   As String : lager   = ws.Cells(zeile, INV_COL_LAGER).Value
    Dim ekS     As Double : ekS     = Val(ws.Cells(zeile, INV_COL_EK).Value)
    Dim attr    As String : attr    = CStr(ws.Cells(zeile, INV_COL_ATTR).Value)
    If ean = "" And artikel = "" Then Exit Sub

    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    Dim ek As Double : ek = 0
    Dim wg As String : wg = ""
    If Not wsA Is Nothing Then
        Dim cEAN2  As Long : cEAN2  = Spalte_Finden(wsA, "EAN13")
        Dim cEK2   As Long : cEK2   = Spalte_Finden(wsA, "EK-PREIS")
        Dim cWG2   As Long : cWG2   = Spalte_Finden(wsA, "WARENGRUPPE")
        Dim cAttr2 As Long : cAttr2 = Spalte_Finden(wsA, "ATTRIBUT")
        Dim j As Long
        For j = ART_DATA_START To wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row
            If CStr(wsA.Cells(j, cEAN2).Value) = ean Then
                If ekS = 0 And cEK2 > 0 Then ekS = Val(wsA.Cells(j, cEK2).Value)
                wg = wsA.Cells(j, cWG2).Value
                If attr = "" And cAttr2 > 0 Then attr = CStr(wsA.Cells(j, cAttr2).Value)
                Exit For
            End If
        Next j
    End If
    InvEingabe_Zeigen artikel, artNr, ean, vk, ekS, lager, wg, attr, soll, zeile
End Sub

' ================================================================
'  INVEINGABE - BLATT ANZEIGEN MIT ARTIKELDATEN
' ================================================================
Sub InvEingabe_Zeigen(artikel As String, artNr As String, ean As String, _
                      vk As Double, ek As Double, lager As String, _
                      wg As String, attr As String, soll As Double, quellZeile As Long)
    Dim wsE As Worksheet : Set wsE = GetSheet("InvEingabe")
    If wsE Is Nothing Then
        MsgBox "InvEingabe-Blatt fehlt. Bitte Setup_InvEingabe ausfuehren.", vbExclamation
        Exit Sub
    End If
    Application.EnableEvents = False
    ' Spalte A leer - Inhalt ab Spalte B/C
    wsE.Cells(3, 3).Value = artikel              ' C3: Artikelname
    wsE.Cells(4, 3).Value = artNr                ' C4: Art.-Nr.
    wsE.Cells(4, 6).NumberFormat = "@"
    wsE.Cells(4, 6).Value = CStr(ean)            ' F4: EAN
    wsE.Cells(5, 3).Value = Format(vk, "0.00") & " EUR"  ' C5: VK-Preis (nur Anzeige)
    wsE.Cells(5, 6).Value = Format(ek, "0.00") & " EUR"  ' F5: EK-Preis (nur Anzeige)
    wsE.Cells(6, 3).Value = lager                ' C6: Lagerort (editierbar)
    wsE.Cells(6, 6).Value = wg                   ' F6: Warengruppe (editierbar)
    wsE.Cells(7, 3).Value = attr                 ' C7: Attribut (editierbar)
    wsE.Cells(9, 3).Value = Format(soll, "0") & " Stk"  ' C9: SOLL-Anzeige
    wsE.Cells(9, 6).Value = ""                   ' F9: GEZAEHLT-Eingabe leeren
    wsE.Cells(10, 2).Value = ""                  ' B10: DIFFERENZ leeren
    wsE.Cells(10, 2).Interior.ColorIndex = xlNone
    wsE.Cells(14, 1).Value = quellZeile          ' A14: Quellzeile (versteckt)
    ' Dropdowns befüllen
    InvEingabe_SetzeDropdowns wsE
    Application.EnableEvents = True
    wsE.Visible = xlSheetVisible
    wsE.Activate
    wsE.Cells(9, 6).Select                       ' Cursor auf GEZAEHLT-Feld
End Sub

' ================================================================
'  INVEINGABE - DROPDOWNS BEFUELLEN (Lagerort + Warengruppe)
'  Liest alle einzigartigen Werte aus dem Artikel-Sheet
' ================================================================
Private Sub InvEingabe_SetzeDropdowns(wsE As Worksheet)
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Dim cLag  As Long : cLag  = Spalte_Finden(wsA, "LAGERORT")
    Dim cWG   As Long : cWG   = Spalte_Finden(wsA, "WARENGRUPPE")
    Dim cAttr As Long : cAttr = Spalte_Finden(wsA, "ATTRIBUT")
    Dim lastRow As Long : lastRow = wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row
    Dim dictLag  As Object : Set dictLag  = CreateObject("Scripting.Dictionary")
    Dim dictWG   As Object : Set dictWG   = CreateObject("Scripting.Dictionary")
    Dim dictAttr As Object : Set dictAttr = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = ART_DATA_START To lastRow
        If cLag > 0 Then
            Dim lv As String : lv = Trim(CStr(wsA.Cells(i, cLag).Value))
            If lv <> "" Then dictLag(lv) = 1
        End If
        If cWG > 0 Then
            Dim wv As String : wv = Trim(CStr(wsA.Cells(i, cWG).Value))
            If wv <> "" Then dictWG(wv) = 1
        End If
        If cAttr > 0 Then
            Dim av As String : av = Trim(CStr(wsA.Cells(i, cAttr).Value))
            If av <> "" Then dictAttr(av) = 1
        End If
    Next i
    ' Dropdown Lagerort → C6
    If dictLag.Count > 0 Then
        With wsE.Cells(6, 3).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                 Formula1:=Join(dictLag.Keys, ",")
            .ShowError = False
        End With
    End If
    ' Dropdown Warengruppe → F6
    If dictWG.Count > 0 Then
        With wsE.Cells(6, 6).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                 Formula1:=Join(dictWG.Keys, ",")
            .ShowError = False
        End With
    End If
    ' Dropdown Attribut → C7 (liest alle einzigartigen Werte aus Artikel-Sheet)
    If dictAttr.Count > 0 Then
        With wsE.Cells(7, 3).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                 Formula1:=Join(dictAttr.Keys, ",")
            .ShowError = False
        End With
    End If
End Sub

' ================================================================
'  INVEINGABE - HANDLER (Klick auf Buttons Zeile 10)
' ================================================================
Sub InvEingabe_Handler(ByVal Target As Range)
    ' Cursor sichtbar in editierbaren Feldern (F2 = Bearbeitungsmodus)
    If (Target.Row = 6 And (Target.Column = 3 Or Target.Column = 6)) Or _
       (Target.Row = 7 And Target.Column = 3) Or _
       (Target.Row = 9 And Target.Column = 6) Then
        Application.SendKeys "{F2}"
        Exit Sub
    End If

    ' Zeile 12: Buttons  B12=ABBRECHEN  E12=UEBERNEHMEN
    If Target.Row = 12 Then
        Select Case Target.Column
            Case 2, 3       : InvEingabe_Abbrechen
            Case 5, 6, 7    : InvEingabe_Uebernehmen
        End Select
        Exit Sub
    End If
    ' Zeile 10 (DIFFERENZ-Anzeige): Enter aus F9 landet hier → auto Uebernehmen
    If Target.Row = 10 Then
        Dim wsE As Worksheet : Set wsE = Target.Worksheet
        If Trim(CStr(wsE.Cells(9, 6).Value)) <> "" Then
            InvEingabe_Uebernehmen
        End If
    End If
End Sub

' ================================================================
'  INVEINGABE - GEZÄHLT EINGABE → DIFFERENZ LIVE ANZEIGEN
' ================================================================
Sub InvEingabe_GezaeltChange(ByVal Target As Range)
    ' GEZAEHLT-Eingabe ist F9 (Zeile 9, Spalte 6)
    If Target.Row <> 9 Or Target.Column <> 6 Then Exit Sub
    Dim wsE   As Worksheet : Set wsE = Target.Worksheet
    Dim soll  As Double    : soll  = Val(wsE.Cells(9, 3).Value)  ' C9
    Dim gezStr As String   : gezStr = Trim(CStr(Target.Value))
    Application.EnableEvents = False
    If gezStr = "" Then
        wsE.Cells(10, 2).Value = ""                              ' B10
        wsE.Cells(10, 2).Interior.ColorIndex = xlNone
    Else
        Dim ist  As Double : ist  = Val(gezStr)
        Dim diff As Double : diff = ist - soll
        Select Case True
            Case diff = 0
                wsE.Cells(10, 2).Value = "Bestand stimmt  +/- 0 Stk"
                wsE.Cells(10, 2).Interior.Color = RGB(198, 239, 206)
                wsE.Cells(10, 2).Font.Color = RGB(0, 97, 0)
            Case diff > 0
                wsE.Cells(10, 2).Value = "+" & Format(diff, "0") & " Stk  (mehr gezaehlt)"
                wsE.Cells(10, 2).Interior.Color = RGB(255, 235, 156)
                wsE.Cells(10, 2).Font.Color = RGB(120, 80, 0)
            Case Else
                wsE.Cells(10, 2).Value = Format(diff, "0") & " Stk  (weniger gezaehlt)"
                wsE.Cells(10, 2).Interior.Color = RGB(255, 199, 206)
                wsE.Cells(10, 2).Font.Color = RGB(156, 0, 6)
        End Select
    End If
    Application.EnableEvents = True
End Sub

' ================================================================
'  INVEINGABE - ABBRECHEN
' ================================================================
Sub InvEingabe_Abbrechen()
    Dim wsE As Worksheet : Set wsE = GetSheet("InvEingabe")
    If Not wsE Is Nothing Then wsE.Visible = xlSheetHidden
    GetSheet("InvSuch").Activate
End Sub

' ================================================================
'  INVEINGABE - ÜBERNEHMEN
' ================================================================
Sub InvEingabe_Uebernehmen()
    Dim wsE As Worksheet : Set wsE = GetSheet("InvEingabe")
    Dim wsS As Worksheet : Set wsS = GetSheet("InvSuch")
    If wsE Is Nothing Or wsS Is Nothing Then Exit Sub
    Dim gezStr As String : gezStr = Trim(CStr(wsE.Cells(9, 6).Value))  ' F9
    If gezStr = "" Then
        MsgBox "Bitte zuerst die gezaehlte Menge eingeben.", vbExclamation : Exit Sub
    End If
    Dim ist        As Double : ist        = Val(gezStr)
    Dim soll       As Double : soll       = Val(wsE.Cells(9, 3).Value)  ' C9
    Dim diff       As Double : diff       = ist - soll
    Dim quellZeile As Long   : quellZeile = Val(wsE.Cells(14, 1).Value) ' A14
    If quellZeile < 4 Then MsgBox "Ungueltige Quellzeile.", vbExclamation : Exit Sub
    On Error GoTo UebernahmeErr
    Application.EnableEvents = False
    wsS.Cells(quellZeile, INV_COL_IST).Value = ist
    wsS.Cells(quellZeile, INV_COL_DIFF).Value = diff
    wsS.Cells(quellZeile, INV_COL_DIFF).NumberFormat = IIf(diff >= 0, "+0;-0;0", "0")
    Select Case True
        Case diff = 0 : wsS.Rows(quellZeile).Interior.Color = RGB(198, 239, 206)
        Case diff > 0 : wsS.Rows(quellZeile).Interior.Color = RGB(255, 235, 156)
        Case Else     : wsS.Rows(quellZeile).Interior.Color = RGB(255, 199, 206)
    End Select
    Application.EnableEvents = True

    ' ── Felder aus editierbaren Feldern lesen ───────────────────
    Dim eanE     As String : eanE     = CStr(wsE.Cells(4, 6).Value)          ' F4
    Dim artNrE   As String : artNrE   = CStr(wsE.Cells(4, 3).Value)          ' C4
    Dim artikelE As String : artikelE = CStr(wsE.Cells(3, 3).Value)          ' C3
    Dim lagerE   As String : lagerE   = Trim(CStr(wsE.Cells(6, 3).Value))    ' C6
    Dim wgE      As String : wgE      = Trim(CStr(wsE.Cells(6, 6).Value))    ' F6
    Dim attrE    As String : attrE    = Trim(CStr(wsE.Cells(7, 3).Value))    ' C7

    ' ── Felder im Artikel-Sheet aktualisieren ───────────────────
    Dim wsA2 As Worksheet : Set wsA2 = GetSheet("Artikel")
    If Not wsA2 Is Nothing Then
        Dim cEANA  As Long : cEANA  = Spalte_Finden(wsA2, "EAN13")
        Dim cLagA  As Long : cLagA  = Spalte_Finden(wsA2, "LAGERORT")
        Dim cWGA   As Long : cWGA   = Spalte_Finden(wsA2, "WARENGRUPPE")
        Dim cAttrA As Long : cAttrA = Spalte_Finden(wsA2, "ATTRIBUT")
        If cEANA = 0 Then GoTo SkipArtikelUpdate
        Dim jA As Long
        For jA = ART_DATA_START To wsA2.Cells(wsA2.Rows.Count, 2).End(xlUp).Row
            If CStr(wsA2.Cells(jA, cEANA).Value) = eanE Then
                If cLagA  > 0 And lagerE <> "" Then wsA2.Cells(jA, cLagA).Value  = lagerE
                If cWGA   > 0 And wgE    <> "" Then wsA2.Cells(jA, cWGA).Value   = wgE
                If cAttrA > 0                  Then wsA2.Cells(jA, cAttrA).Value  = attrE
                ' Attribut auch in InvSuche-Zeile aktualisieren
                wsS.Cells(quellZeile, INV_COL_ATTR).Value = attrE
                Exit For
            End If
        Next jA
    End If

SkipArtikelUpdate:
    ' ── Persistent in InvDaten speichern ────────────────────────
    InvDaten_Eintragen eanE, artNrE, artikelE, soll, ist, diff, lagerE

    wsE.Visible = xlSheetHidden
    wsS.Activate
    Exit Sub

UebernahmeErr:
    Application.EnableEvents = True
    MsgBox "Fehler beim Uebernehmen: " & Err.Description & " (Nr. " & Err.Number & ")", _
           vbExclamation, "Fehler"
End Sub

' ================================================================
'  INVDATEN - STARTDATUM LESEN
' ================================================================
Function InvDaten_StartDatum() As Date
    Dim wsD As Worksheet : Set wsD = GetSheet("InvDaten")
    If wsD Is Nothing Then InvDaten_StartDatum = DateSerial(2000, 1, 1) : Exit Function
    Dim d As Variant : d = wsD.Cells(1, 2).Value
    If IsDate(d) Then InvDaten_StartDatum = CDate(d) Else InvDaten_StartDatum = DateSerial(2000, 1, 1)
End Function

' ================================================================
'  INVDATEN - IST-WERT FÜR EAN HOLEN (neuester Eintrag ab Startdatum)
' ================================================================
Function InvDaten_IstHolen(ean As String) As String
    Dim wsD As Worksheet : Set wsD = GetSheet("InvDaten")
    If wsD Is Nothing Then InvDaten_IstHolen = "" : Exit Function
    If ean = "" Then InvDaten_IstHolen = "" : Exit Function
    Dim startD As Date : startD = InvDaten_StartDatum()
    Dim lastRow As Long : lastRow = wsD.Cells(wsD.Rows.Count, 2).End(xlUp).Row
    Dim i As Long, result As String : result = ""
    For i = 4 To lastRow
        If CStr(wsD.Cells(i, 2).Value) = ean Then
            Dim entryDate As Variant : entryDate = wsD.Cells(i, 1).Value
            If IsDate(entryDate) Then
                If CDate(entryDate) >= startD Then
                    result = CStr(wsD.Cells(i, 6).Value) ' neuester gewinnt
                End If
            End If
        End If
    Next i
    InvDaten_IstHolen = result
End Function

' ================================================================
'  INVDATEN - EINTRAG SCHREIBEN (bei jedem UEBERNEHMEN)
' ================================================================
Sub InvDaten_Eintragen(ean As String, artNr As String, artikel As String, _
                       soll As Double, ist As Double, diff As Double, lager As String)
    Dim wsD As Worksheet : Set wsD = GetSheet("InvDaten")
    If wsD Is Nothing Then Exit Sub
    Dim nRow As Long : nRow = wsD.Cells(wsD.Rows.Count, 2).End(xlUp).Row + 1
    If nRow < 4 Then nRow = 4
    wsD.Cells(nRow, 1).Value = Now()
    wsD.Cells(nRow, 1).NumberFormat = "DD.MM.YYYY HH:MM"
    wsD.Cells(nRow, 2).NumberFormat = "@" : wsD.Cells(nRow, 2).Value = CStr(ean)
    wsD.Cells(nRow, 3).NumberFormat = "@" : wsD.Cells(nRow, 3).Value = CStr(artNr)
    wsD.Cells(nRow, 4).Value = artikel
    wsD.Cells(nRow, 5).Value = soll
    wsD.Cells(nRow, 6).Value = ist
    wsD.Cells(nRow, 7).Value = diff
    wsD.Cells(nRow, 8).Value = lager
End Sub

' ================================================================
'  INVENTUR STARTEN - Startdatum festlegen
' ================================================================
Sub Inventur_Starten()
    Dim wsD As Worksheet : Set wsD = GetSheet("InvDaten")
    If wsD Is Nothing Then
        MsgBox "InvDaten-Sheet fehlt. Bitte Setup_InvDaten ausfuehren.", vbExclamation : Exit Sub
    End If
    Dim eingabe As String
    eingabe = InputBox("Inventur-Startdatum eingeben:" & Chr(10) & _
                       "(leer = heute)", "Inventur starten", Format(Date, "DD.MM.YYYY"))
    If eingabe = "" Then Exit Sub
    If Not IsDate(eingabe) Then MsgBox "Ungültiges Datum.", vbExclamation : Exit Sub
    Dim startD As Date : startD = CDate(eingabe)
    wsD.Cells(1, 2).Value = startD
    wsD.Cells(1, 2).NumberFormat = "DD.MM.YYYY"
    MsgBox "Inventur gestartet ab: " & Format(startD, "DD.MM.YYYY") & Chr(10) & Chr(10) & _
           "Nur Zählungen ab diesem Datum werden angezeigt.", _
           vbInformation, "Inventur gestartet"
End Sub

' ================================================================
'  GEZAEHLT ANZEIGEN - alle Artikel seit Startdatum in InvSuche
' ================================================================
Sub InvDaten_GezaehltAnzeigen()
    Dim wsD As Worksheet : Set wsD = GetSheet("InvDaten")
    Dim wsS As Worksheet : Set wsS = GetSheet("InvSuch")
    If wsD Is Nothing Or wsS Is Nothing Then Exit Sub

    Dim startD As Date : startD = InvDaten_StartDatum()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Alte Ergebnisse löschen
    Dim lastOld As Long : lastOld = wsS.Cells(wsS.Rows.Count, INV_COL_ARTIKEL).End(xlUp).Row
    If lastOld >= 4 Then
        wsS.Rows("4:" & lastOld).ClearContents
        wsS.Rows("4:" & lastOld).Interior.ColorIndex = xlNone
    End If
    InvSuche_Headers wsS
    wsS.Cells(2, 3).Value = "Alle gezaehlt seit " & Format(startD, "DD.MM.YYYY")

    ' Neuesten Eintrag pro EAN sammeln
    Dim dictIst  As Object : Set dictIst  = CreateObject("Scripting.Dictionary")
    Dim dictInfo As Object : Set dictInfo = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long : lastRow = wsD.Cells(wsD.Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    For i = 4 To lastRow
        Dim entryDate As Variant : entryDate = wsD.Cells(i, 1).Value
        If IsDate(entryDate) Then
            If CDate(entryDate) >= startD Then
                Dim eKey As String : eKey = CStr(wsD.Cells(i, 2).Value)
                dictIst(eKey)  = CStr(wsD.Cells(i, 6).Value)
                dictInfo(eKey) = CStr(wsD.Cells(i, 3).Value) & "|" & _
                                 CStr(wsD.Cells(i, 4).Value) & "|" & _
                                 CStr(wsD.Cells(i, 5).Value) & "|" & _
                                 CStr(wsD.Cells(i, 7).Value) & "|" & _
                                 CStr(wsD.Cells(i, 8).Value)
            End If
        End If
    Next i

    ' Zeilen schreiben
    Dim sRow As Long : sRow = 4
    Dim key As Variant
    For Each key In dictIst.Keys
        Dim info() As String : info = Split(CStr(dictInfo(key)), "|")
        Dim istV  As Double : istV  = Val(CStr(dictIst(key)))
        Dim sollV As Double : sollV = Val(info(2))
        Dim diffV As Double : diffV = istV - sollV
        wsS.Cells(sRow, INV_COL_NR).Value = sRow - 3
        wsS.Cells(sRow, INV_COL_ARTNR).NumberFormat = "@"
        wsS.Cells(sRow, INV_COL_ARTNR).Value = info(0)
        wsS.Cells(sRow, INV_COL_ARTIKEL).Value = info(1)
        wsS.Cells(sRow, INV_COL_SOLL).Value = sollV
        wsS.Cells(sRow, INV_COL_IST).Value = istV
        wsS.Cells(sRow, INV_COL_DIFF).Value = diffV
        wsS.Cells(sRow, INV_COL_DIFF).NumberFormat = IIf(diffV >= 0, "+0;-0;0", "0")
        wsS.Cells(sRow, INV_COL_EAN).NumberFormat = "@"
        wsS.Cells(sRow, INV_COL_EAN).Value = CStr(key)
        wsS.Cells(sRow, INV_COL_LAGER).Value = info(4)
        Select Case True
            Case diffV = 0 : wsS.Rows(sRow).Interior.Color = RGB(198, 239, 206)
            Case diffV > 0 : wsS.Rows(sRow).Interior.Color = RGB(255, 235, 156)
            Case Else      : wsS.Rows(sRow).Interior.Color = RGB(255, 199, 206)
        End Select
        sRow = sRow + 1
    Next key

    wsS.Cells(2, 10).Value = (sRow - 4) & " gezaehlt"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    wsS.Activate
End Sub

' ================================================================
'  SETUP: InvDaten-Blatt erstellen (einmalig ausführen)
' ================================================================
Sub Setup_InvDaten()
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "InvDaten", vbTextCompare) > 0 Then ws.Delete : Exit For
    Next ws
    Application.DisplayAlerts = True

    Dim wsD As Worksheet
    Set wsD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsD.Name = "InvDaten"

    ' Startdatum in A1/B1
    wsD.Cells(1, 1).Value = "Inventur-Start:"
    wsD.Cells(1, 1).Font.Bold = True
    wsD.Cells(1, 2).Value = Date
    wsD.Cells(1, 2).NumberFormat = "DD.MM.YYYY"
    wsD.Columns("A").ColumnWidth = 18

    ' Header Zeile 3
    Dim hdr As Variant
    hdr = Array("Datum/Zeit", "EAN", "Art.-Nr.", "Artikel", "SOLL", "IST", "DIFFERENZ", "Lagerort")
    Dim col As Long
    For col = 1 To 8
        With wsD.Cells(3, col)
            .Value = hdr(col - 1)
            .Font.Bold = True
            .Interior.Color = RGB(64, 64, 64)
            .Font.Color = RGB(255, 255, 255)
        End With
    Next col
    wsD.Columns("B").ColumnWidth = 16
    wsD.Columns("C").ColumnWidth = 14
    wsD.Columns("D").ColumnWidth = 30
    wsD.Columns("E:G").ColumnWidth = 8
    wsD.Columns("H").ColumnWidth = 14

    wsD.Visible = xlSheetHidden
    MsgBox "InvDaten-Sheet erstellt!" & Chr(10) & _
           "Startdatum: " & Format(Date, "DD.MM.YYYY") & Chr(10) & Chr(10) & _
           "Mit 'INVENTUR STARTEN' (Button in InvSuche) kannst du das Datum anpassen.", _
           vbInformation, "Setup_InvDaten"
End Sub

' ================================================================
'  SETUP: InvSuche-Blatt formatieren (einmalig ausführen)
'  Passt Layout, Spaltenbreiten, Header und Buttons an.
'  Vorhandene Daten (ab Zeile 4) bleiben erhalten.
' ================================================================
Sub Setup_InvSuche()
    Dim wsS As Worksheet : Set wsS = GetSheet("InvSuch")
    If wsS Is Nothing Then
        MsgBox "InvSuche-Blatt nicht gefunden!" & Chr(10) & SheetListe(), vbCritical : Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ── Spaltenbreiten ──────────────────────────────────────────
    wsS.Columns("A").ColumnWidth = 3     ' leer (Rand)
    wsS.Columns("B").ColumnWidth = 5     ' Nr
    wsS.Columns("C").ColumnWidth = 14    ' Art.-Nr.
    wsS.Columns("D").ColumnWidth = 32    ' Artikel
    wsS.Columns("E").ColumnWidth = 8     ' SOLL
    wsS.Columns("F").ColumnWidth = 11    ' IST / Gezählt
    wsS.Columns("G").ColumnWidth = 11    ' DIFFERENZ
    wsS.Columns("H").ColumnWidth = 12    ' VK-Preis / SUCHEN
    wsS.Columns("I").ColumnWidth = 12    ' EK-Preis / LEEREN
    wsS.Columns("J").ColumnWidth = 15    ' EAN
    wsS.Columns("K").ColumnWidth = 14    ' Lagerort / UEBERNEHMEN
    wsS.Columns("L").ColumnWidth = 16    ' Attribut

    ' ── Zeilen einfrieren (1-3 immer sichtbar) ──────────────────
    wsS.Activate
    wsS.Cells(4, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ' ── Zeilenhöhen ─────────────────────────────────────────────
    wsS.Rows(1).RowHeight = 28
    wsS.Rows(2).RowHeight = 24
    wsS.Rows(3).RowHeight = 22

    ' ── Zeile 1: Titelzeile (A1:K1) ─────────────────────────────
    On Error Resume Next : wsS.Range("A1:K1").UnMerge : On Error GoTo 0
    wsS.Range("A1:K1").Merge
    With wsS.Cells(1, 1)
        .Value = "Inventur Artikelsuche"
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' ── Zeile 2: Suchfeld (C2:G2) + Buttons (H2, I2, K2) ───────
    ' A2: leer  B2: "Suche:" Label
    wsS.Cells(2, 1).Value = ""
    wsS.Cells(2, 2).Value = "Suche:"
    wsS.Cells(2, 2).Font.Bold = True
    wsS.Cells(2, 2).Font.Name = "Arial"
    wsS.Cells(2, 2).Font.Size = 12

    ' Suchfeld C2:G2 (gelb, umrandet) - Cursor erscheint beim Klick (Sheet-Modul)
    On Error Resume Next : wsS.Range("C2:G2").UnMerge : On Error GoTo 0
    wsS.Range("C2:G2").Merge
    With wsS.Cells(2, 3)
        .Interior.Color = RGB(255, 255, 200)
        .Font.Name = "Arial"
        .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
    End With

    ' H2: SUCHEN (grün)
    With wsS.Cells(2, 8)
        .Value = "SUCHEN"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 120, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' I2: LEEREN (orange)
    With wsS.Cells(2, 9)
        .Value = "LEEREN"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(180, 90, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' J2: leer
    wsS.Cells(2, 10).Value = ""
    wsS.Cells(2, 10).Interior.ColorIndex = xlNone

    ' K2: UEBERNEHMEN (blau)
    With wsS.Cells(2, 11)
        .Value = "UEBERNEHMEN"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 80, 160)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' J2: Treffer-Anzeige (Spalte 10, in Zeile 2 frei)
    With wsS.Cells(2, 10)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.ColorIndex = xlNone
    End With

    ' L2: INVENTUR STARTEN (dunkelgrau)
    With wsS.Cells(2, 12)
        .Value = "INVENTUR STARTEN"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
        .ColumnWidth = 18
    End With

    ' M2: GEZAEHLT ANZEIGEN (dunkelblau)
    With wsS.Cells(2, 13)
        .Value = "GEZAEHLT ANZEIGEN"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 60, 120)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
        .ColumnWidth = 19
    End With

    ' ── Zeile 3: Spaltenüberschriften (alle Spalten korrekt) ────
    InvSuche_Headers wsS

    ' ── Standardformatierung Datenbereich ───────────────────────
    With wsS.Cells
        .Font.Name = "Arial"
        .Font.Size = 11
    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "InvSuche-Layout aktualisiert!" & Chr(10) & Chr(10) & _
           "Zeile 1  = Titelzeile" & Chr(10) & _
           "Zeile 2  = Suchfeld (B2, gelb)" & Chr(10) & _
           "Zeile 3  = Spaltenkoepfe  +  Buttons (G=SUCHEN, H=LEEREN, J=UEBERNEHMEN)" & Chr(10) & _
           "Zeile 4+ = Suchergebnisse" & Chr(10) & Chr(10) & _
           "Bitte Strg+S speichern.", _
           vbInformation, "Setup_InvSuche abgeschlossen"
End Sub

' ================================================================
'  SETUP: InvEingabe-Blatt erstellen (einmalig ausführen)
' ================================================================
Sub Setup_InvEingabe()
    ' Sheet NICHT loeschen - nur neu formatieren, damit Sheet-Modul-Code erhalten bleibt
    Dim wsE As Worksheet
    Set wsE = GetSheet("InvEingabe")
    If wsE Is Nothing Then
        ' Nur beim allerersten Mal neu anlegen
        Set wsE = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsE.Name = "InvEingabe"
        MsgBox "Neues InvEingabe-Blatt angelegt." & Chr(10) & _
               "Bitte jetzt den Sheet-Modul-Code einfuegen (Alt+F11 -> InvEingabe).", _
               vbInformation, "Hinweis"
    End If

    ' Alles loeschen (Inhalt + Formatierung + Merges + Validierungen)
    Application.DisplayAlerts = False
    wsE.Visible = xlSheetVisible
    wsE.Cells.UnMerge
    wsE.Cells.ClearContents
    wsE.Cells.ClearFormats
    wsE.Cells.Validation.Delete
    Application.DisplayAlerts = True

    wsE.Cells.Font.Name = "Arial"
    wsE.Cells.Font.Size = 12
    wsE.Cells.RowHeight = 22

    ' Spaltenbreiten: A=leer(Rand), B=Label, C=Wert, D=leer, E=Label, F=Wert, G=Rand
    wsE.Columns("A").ColumnWidth = 3
    wsE.Columns("B").ColumnWidth = 14
    wsE.Columns("C").ColumnWidth = 22
    wsE.Columns("D").ColumnWidth = 2
    wsE.Columns("E").ColumnWidth = 14
    wsE.Columns("F").ColumnWidth = 22
    wsE.Columns("G").ColumnWidth = 3

    ' ── Zeile 1: Titelzeile ─────────────────────────────────────
    wsE.Rows(1).RowHeight = 30
    wsE.Range("A1:G1").Merge
    With wsE.Cells(1, 1)
        .Value = "Inventur - Artikel pruefen"
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 14 : .Font.Bold = True
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With

    ' ── Zeile 3: Artikelname (Zeilenumbruch aktiviert) ──────────
    wsE.Rows(3).RowHeight = 40
    wsE.Cells(3, 2).Value = "Artikel:" : wsE.Cells(3, 2).Font.Bold = True
    wsE.Range("C3:G3").Merge
    With wsE.Cells(3, 3)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    ' ── Zeile 4: Art.-Nr. + EAN ─────────────────────────────────
    wsE.Cells(4, 2).Value = "Art.-Nr.:" : wsE.Cells(4, 2).Font.Bold = True
    wsE.Cells(4, 5).Value = "EAN:" : wsE.Cells(4, 5).Font.Bold = True
    wsE.Range("F4:G4").Merge
    wsE.Cells(4, 6).NumberFormat = "@"

    ' ── Zeile 5: VK-Preis + EK-Preis (nur Anzeige, grau) ───────
    wsE.Cells(5, 2).Value = "VK-Preis:" : wsE.Cells(5, 2).Font.Bold = True
    With wsE.Cells(5, 3)
        .Interior.Color = RGB(240, 240, 240)
        .Font.Color = RGB(80, 80, 80)
    End With
    wsE.Cells(5, 5).Value = "EK-Preis:" : wsE.Cells(5, 5).Font.Bold = True
    wsE.Range("F5:G5").Merge
    With wsE.Cells(5, 6)
        .Interior.Color = RGB(240, 240, 240)
        .Font.Color = RGB(80, 80, 80)
    End With

    ' ── Zeile 6: Lagerort + Warengruppe (editierbar) ────────────
    wsE.Cells(6, 2).Value = "Lagerort:" : wsE.Cells(6, 2).Font.Bold = True
    With wsE.Cells(6, 3)                          ' C6: Lagerort
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsE.Cells(6, 5).Value = "Warengruppe:" : wsE.Cells(6, 5).Font.Bold = True
    wsE.Range("F6:G6").Merge
    With wsE.Cells(6, 6)                          ' F6: Warengruppe
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With

    ' ── Zeile 7: Attribut (editierbar, Auswahlliste) ────────────
    wsE.Cells(7, 2).Value = "Attribut:" : wsE.Cells(7, 2).Font.Bold = True
    wsE.Range("C7:G7").Merge
    With wsE.Cells(7, 3)                          ' C7: Attribut
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With

    ' ── Zeile 8: Trennlinie ─────────────────────────────────────
    wsE.Rows(8).RowHeight = 4
    wsE.Range("A8:G8").Interior.Color = RGB(180, 180, 180)

    ' ── Zeile 9: SOLL + GEZAEHLT ────────────────────────────────
    wsE.Rows(9).RowHeight = 28
    wsE.Cells(9, 2).Value = "SOLL:" : wsE.Cells(9, 2).Font.Bold = True
    wsE.Cells(9, 3).Font.Bold = True
    wsE.Cells(9, 3).HorizontalAlignment = xlCenter
    wsE.Cells(9, 5).Value = "GEZAEHLT:" : wsE.Cells(9, 5).Font.Bold = True
    wsE.Range("F9:G9").Merge
    With wsE.Cells(9, 6)                          ' F9: GEZAEHLT-Eingabe
        .Interior.Color = RGB(255, 255, 200)
        .Font.Bold = True : .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' ── Zeile 10: DIFFERENZ-Anzeige ─────────────────────────────
    wsE.Rows(10).RowHeight = 26
    wsE.Range("B10:G10").Merge
    wsE.Cells(10, 2).Font.Bold = True
    wsE.Cells(10, 2).HorizontalAlignment = xlCenter
    wsE.Cells(10, 2).VerticalAlignment = xlCenter

    ' ── Zeile 12: Buttons ───────────────────────────────────────
    wsE.Rows(12).RowHeight = 28
    With wsE.Cells(12, 2)                         ' B12: ABBRECHEN
        .Value = "ABBRECHEN"
        .Font.Bold = True : .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With
    wsE.Range("E12:G12").Merge
    With wsE.Cells(12, 5)                         ' E12: UEBERNEHMEN
        .Value = "UEBERNEHMEN"
        .Font.Bold = True : .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(64, 64, 64)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' ── Zeile 14: Quellzeile (versteckt in A) ───────────────────
    wsE.Rows(14).RowHeight = 4
    wsE.Cells(14, 1).Font.Color = RGB(255, 255, 255)

    wsE.Visible = xlSheetHidden

    MsgBox "InvEingabe-Blatt neu erstellt!" & Chr(10) & Chr(10) & _
           "Zeile 5  = VK/EK-Preis (Anzeige)" & Chr(10) & _
           "Zeile 6  = Lagerort + Warengruppe (editierbar, Dropdown)" & Chr(10) & _
           "Zeile 7  = Attribut (editierbar, Dropdown aus Artikel-Sheet)" & Chr(10) & _
           "Zeile 9  = SOLL + GEZAEHLT" & Chr(10) & _
           "Zeile 12 = Buttons" & Chr(10) & Chr(10) & _
           "Strg+S speichern.", vbInformation, "Setup abgeschlossen"
End Sub

' ================================================================
'  SETUP: InvArtikelForm automatisch erstellen (einmalig ausfuehren)
'  Voraussetzung: Trust Center -> Makroeinstellungen ->
'  Haken bei "Zugriff auf das VBA-Projektobjektmodell vertrauen"
'  HINWEIS: Wird bei der neuen Inline-Bearbeitung nicht mehr
'  benoetigt, bleibt aber als Fallback erhalten.
' ================================================================
Sub Setup_InvArtikelForm()
    Dim vbp As Object
    On Error GoTo NeedTrust
    Set vbp = ThisWorkbook.VBProject
    On Error GoTo 0

    ' Alte Form loeschen falls vorhanden
    Dim comp As Object
    For Each comp In vbp.VBComponents
        If comp.Name = "InvArtikelForm" Then
            vbp.VBComponents.Remove comp
            Exit For
        End If
    Next comp

    ' Neue UserForm anlegen
    Dim frm As Object
    Set frm = vbp.VBComponents.Add(3)
    frm.Name = "InvArtikelForm"
    With frm.Properties
        .Item("Caption")         = "Inventur - Artikel pruefen"
        .Item("Width")           = 490
        .Item("Height")          = 230
        .Item("StartUpPosition") = 1
        .Item("BackColor")       = RGB(255, 255, 255)
    End With

    Dim d As Object : Set d = frm.Designer
    Dim c As Object

    ' --- Header (dunkelgrau) ---
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_Header"
    c.Caption = "Inventur - Artikel pruefen"
    c.Left = 0 : c.Top = 0 : c.Width = 488 : c.Height = 28
    c.BackColor = RGB(64, 64, 64) : c.ForeColor = RGB(255, 255, 255)
    c.TextAlign = 2 : c.Font.Name = "Arial" : c.Font.Size = 14 : c.Font.Bold = True

    ' --- Artikelzeile ---
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_key1"
    c.Caption = "Artikel:" : c.Left = 6 : c.Top = 36 : c.Width = 66 : c.Height = 18
    c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_Artikel"
    c.Caption = "" : c.Left = 76 : c.Top = 36 : c.Width = 406 : c.Height = 18
    c.Font.Size = 12 : c.Font.Name = "Arial"

    ' --- ArtNr + EAN ---
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_key2"
    c.Caption = "Art.-Nr.:" : c.Left = 6 : c.Top = 58 : c.Width = 66 : c.Height = 18
    c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_ArtNr"
    c.Caption = "" : c.Left = 76 : c.Top = 58 : c.Width = 150 : c.Height = 18
    c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_key3"
    c.Caption = "EAN:" : c.Left = 240 : c.Top = 58 : c.Width = 42 : c.Height = 18
    c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_EAN"
    c.Caption = "" : c.Left = 286 : c.Top = 58 : c.Width = 196 : c.Height = 18
    c.Font.Size = 12 : c.Font.Name = "Arial"

    ' --- Lagerort + WG ---
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_key6"
    c.Caption = "Lagerort:" : c.Left = 6 : c.Top = 80 : c.Width = 66 : c.Height = 18
    c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_Lager"
    c.Caption = "" : c.Left = 76 : c.Top = 80 : c.Width = 150 : c.Height = 18
    c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_key7"
    c.Caption = "Warengruppe:" : c.Left = 240 : c.Top = 80 : c.Width = 96 : c.Height = 18
    c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_WG"
    c.Caption = "" : c.Left = 340 : c.Top = 80 : c.Width = 142 : c.Height = 18
    c.Font.Size = 12 : c.Font.Name = "Arial"

    ' --- Trennlinie ---
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_sep"
    c.Caption = "" : c.Left = 0 : c.Top = 104 : c.Width = 488 : c.Height = 2
    c.BackColor = RGB(180, 180, 180)

    ' --- SOLL + GEZÄHLT in einer Reihe ---
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_keySoll"
    c.Caption = "SOLL:" : c.Left = 6 : c.Top = 112 : c.Width = 46 : c.Height = 20
    c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_Soll"
    c.Caption = "" : c.Left = 56 : c.Top = 110 : c.Width = 80 : c.Height = 22
    c.TextAlign = 1 : c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"

    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_keyIst"
    c.Caption = "GEZAEHLT:" : c.Left = 150 : c.Top = 112 : c.Width = 80 : c.Height = 20
    c.ForeColor = RGB(64, 64, 64) : c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"
    Set c = d.Controls.Add("Forms.TextBox.1") : c.Name = "txt_Ist"
    c.Left = 234 : c.Top = 108 : c.Width = 100 : c.Height = 26
    c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"

    ' --- Differenz ---
    Set c = d.Controls.Add("Forms.Label.1") : c.Name = "lbl_Diff"
    c.Caption = "" : c.Left = 0 : c.Top = 142 : c.Width = 488 : c.Height = 22
    c.TextAlign = 2 : c.Font.Bold = True : c.Font.Size = 12 : c.Font.Name = "Arial"

    ' --- Buttons ---
    Set c = d.Controls.Add("Forms.CommandButton.1") : c.Name = "btn_Abbrechen"
    c.Caption = "Abbrechen" : c.Left = 280 : c.Top = 174 : c.Width = 90 : c.Height = 26
    c.Font.Name = "Arial" : c.Font.Size = 12
    Set c = d.Controls.Add("Forms.CommandButton.1") : c.Name = "btn_OK"
    c.Caption = "Uebernehmen" : c.Left = 382 : c.Top = 174 : c.Width = 100 : c.Height = 26
    c.BackColor = RGB(64, 64, 64) : c.ForeColor = RGB(255, 255, 255)
    c.Font.Bold = True : c.Font.Name = "Arial" : c.Font.Size = 12

    Dim code As String
    code = "Public Sub Befuellen(artikel As String, artNr As String, ean As String, _" & Chr(10)
    code = code & "                     vk As Double, ek As Double, lager As String, _" & Chr(10)
    code = code & "                     wg As String, soll As Double, quellZeile As Long)" & Chr(10)
    code = code & "    Me.Tag = CStr(quellZeile)" & Chr(10)
    code = code & "    lbl_Artikel.Caption = artikel" & Chr(10)
    code = code & "    lbl_ArtNr.Caption = artNr" & Chr(10)
    code = code & "    lbl_EAN.Caption = ean" & Chr(10)
    code = code & "    lbl_Lager.Caption = lager" & Chr(10)
    code = code & "    lbl_WG.Caption = wg" & Chr(10)
    code = code & "    lbl_Soll.Caption = Format(soll, ""0"") & "" Stk""" & Chr(10)
    code = code & "    txt_Ist.Value = """"" & Chr(10)
    code = code & "    lbl_Diff.Caption = """" : lbl_Diff.BackColor = &H8000000F" & Chr(10)
    code = code & "    txt_Ist.SetFocus" & Chr(10)
    code = code & "End Sub" & Chr(10) & Chr(10)
    code = code & "Private Sub txt_Ist_Change()" & Chr(10)
    code = code & "    If Trim(txt_Ist.Value) = """" Then lbl_Diff.Caption = """" : lbl_Diff.BackColor = &H8000000F : Exit Sub" & Chr(10)
    code = code & "    Dim ist As Double : ist = Val(txt_Ist.Value)" & Chr(10)
    code = code & "    Dim soll As Double : soll = Val(Replace(lbl_Soll.Caption, "" Stk"", """"))" & Chr(10)
    code = code & "    Dim diff As Double : diff = ist - soll" & Chr(10)
    code = code & "    If diff > 0 Then" & Chr(10)
    code = code & "        lbl_Diff.Caption = ""+"" & Format(diff,""0"") & "" Stk  (mehr vorhanden)""" & Chr(10)
    code = code & "        lbl_Diff.BackColor = RGB(198,239,206) : lbl_Diff.ForeColor = RGB(0,97,0)" & Chr(10)
    code = code & "    ElseIf diff < 0 Then" & Chr(10)
    code = code & "        lbl_Diff.Caption = Format(diff,""0"") & "" Stk  (weniger vorhanden)""" & Chr(10)
    code = code & "        lbl_Diff.BackColor = RGB(255,199,206) : lbl_Diff.ForeColor = RGB(156,0,6)" & Chr(10)
    code = code & "    Else" & Chr(10)
    code = code & "        lbl_Diff.Caption = ""+/-0  Bestand stimmt""" & Chr(10)
    code = code & "        lbl_Diff.BackColor = RGB(198,239,206) : lbl_Diff.ForeColor = RGB(0,97,0)" & Chr(10)
    code = code & "    End If" & Chr(10)
    code = code & "End Sub" & Chr(10) & Chr(10)
    code = code & "Private Sub txt_Ist_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & Chr(10)
    code = code & "    If KeyCode = 13 Then btn_OK_Click" & Chr(10)
    code = code & "    If KeyCode = 27 Then btn_Abbrechen_Click" & Chr(10)
    code = code & "End Sub" & Chr(10) & Chr(10)
    code = code & "Private Sub btn_OK_Click()" & Chr(10)
    code = code & "    If Trim(txt_Ist.Value) = """" Then MsgBox ""Bitte Menge eingeben."", vbExclamation : Exit Sub" & Chr(10)
    code = code & "    Dim ist As Double : ist = Val(txt_Ist.Value)" & Chr(10)
    code = code & "    Dim soll As Double : soll = Val(Replace(lbl_Soll.Caption, "" Stk"", """"))" & Chr(10)
    code = code & "    Dim diff As Double : diff = ist - soll" & Chr(10)
    code = code & "    Dim zeile As Long : zeile = Val(Me.Tag)" & Chr(10)
    code = code & "    Dim ws As Worksheet" & Chr(10)
    code = code & "    For Each ws In ThisWorkbook.Sheets" & Chr(10)
    code = code & "        If InStr(1, ws.Name, ""InvSuch"", vbTextCompare) > 0 Then Exit For" & Chr(10)
    code = code & "    Next ws" & Chr(10)
    code = code & "    If Not ws Is Nothing And zeile > 0 Then" & Chr(10)
    code = code & "        ws.Cells(zeile, 8).Value = ist" & Chr(10)
    code = code & "        If diff = 0 Then ws.Rows(zeile).Interior.Color = RGB(198,239,206)" & Chr(10)
    code = code & "        If diff > 0 Then ws.Rows(zeile).Interior.Color = RGB(255,235,156)" & Chr(10)
    code = code & "        If diff < 0 Then ws.Rows(zeile).Interior.Color = RGB(255,199,206)" & Chr(10)
    code = code & "    End If" & Chr(10)
    code = code & "    Me.Hide" & Chr(10)
    code = code & "End Sub" & Chr(10) & Chr(10)
    code = code & "Private Sub btn_Abbrechen_Click()" & Chr(10)
    code = code & "    Me.Hide" & Chr(10)
    code = code & "End Sub" & Chr(10)

    frm.CodeModule.AddFromString code

    MsgBox "InvArtikelForm wurde erfolgreich erstellt!" & Chr(10) & _
           "Bitte Datei speichern (Strg+S).", vbInformation, "Setup abgeschlossen"
    Exit Sub

NeedTrust:
    MsgBox "Einstellung fehlt:" & Chr(10) & Chr(10) & _
           "Excel -> Datei -> Optionen -> Trust Center ->" & Chr(10) & _
           "Trust Center-Einstellungen -> Makroeinstellungen ->" & Chr(10) & _
           "Haken bei: Zugriff auf das VBA-Projektobjektmodell vertrauen" & Chr(10) & Chr(10) & _
           "Danach Excel neu starten und Setup_InvArtikelForm erneut ausfuehren.", _
           vbExclamation, "Einstellung fehlt"
End Sub

' ================================================================
'  DATENSICHERUNG  -  wird automatisch via Workbook_BeforeSave
'  aufgerufen. Erstellt Kopie im Unterordner \Backup mit Timestamp.
'  Max. BACKUP_MAX Kopien - älteste werden automatisch gelöscht.
'  Einstellung oben: Const BACKUP_MAX = 4
' ================================================================
Sub Datensicherung()
    On Error GoTo SicherungsFehler

    Dim backupDir As String
    backupDir = ThisWorkbook.Path & "\Backup"

    ' Backup-Ordner anlegen falls nicht vorhanden
    If Dir(backupDir, vbDirectory) = "" Then MkDir backupDir

    ' Dateiname mit Datum + Uhrzeit (Minute-genau)
    Dim ts       As String : ts       = Format(Now(), "YYYY-MM-DD_HH-MM")
    Dim origName As String : origName = ThisWorkbook.Name
    Dim destPath As String : destPath = backupDir & "\" & ts & "_" & origName

    ' Diese Minute schon gesichert? → überspringen
    If Dir(destPath) <> "" Then Exit Sub

    ' Aktuelle Version (inkl. ungespeicherter Änderungen) sichern
    ThisWorkbook.SaveCopyAs destPath

    ' ── Alte Backups aufräumen (nur BACKUP_MAX behalten) ─────────
    Dim ext As String : ext = "*" & Right(origName, 5)  ' *.xlsm
    Dim allFiles() As String
    ReDim allFiles(100)
    Dim count As Long : count = 0
    Dim f As String
    f = Dir(backupDir & "\" & ext)
    Do While f <> "" And count <= 100
        allFiles(count) = f : count = count + 1 : f = Dir()
    Loop
    If count <= BACKUP_MAX Then Exit Sub

    ' Bubble-Sort aufsteigend (Timestamp im Namen → älteste zuerst)
    Dim i As Long, j As Long, tmp As String
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If allFiles(j) < allFiles(i) Then
                tmp = allFiles(i) : allFiles(i) = allFiles(j) : allFiles(j) = tmp
            End If
        Next j
    Next i

    ' Älteste löschen bis nur noch BACKUP_MAX übrig
    For i = 0 To count - BACKUP_MAX - 1
        Kill backupDir & "\" & allFiles(i)
    Next i
    Exit Sub

SicherungsFehler:
    ' Fehler still ignorieren - Backup soll den normalen Speichervorgang nie blockieren
End Sub

' ================================================================
'  SCHNELLANSICHT - SYNC-HILFE (EAN → Bestand in Schnellansicht)
' ================================================================
Private Sub Schnellansicht_BestandSync(ean As String, neuerBestand As Double)
    Dim wsSc As Worksheet : Set wsSc = GetSheet("Schnell")
    If wsSc Is Nothing Then Exit Sub
    Dim k As Long
    For k = 4 To wsSc.Cells(wsSc.Rows.Count, SA_COL_EAN).End(xlUp).Row
        If CStr(wsSc.Cells(k, SA_COL_EAN).Value) = ean Then
            wsSc.Cells(k, SA_COL_BEST).Value = neuerBestand : Exit For
        End If
    Next k
End Sub

' ================================================================
'  SCHNELLANSICHT - SPALTENKÖPFE
' ================================================================
Private Sub Schnellansicht_Headers(wsS As Worksheet)
    wsS.Cells(3, 1).Value = "" : wsS.Cells(3, 1).Interior.ColorIndex = xlNone
    Dim headers As Variant
    headers = Array("#", "Art.-Nr.", "Artikel", "EAN", "VK-Preis", _
                    "Bestand", "Einheit", "Lagerort", "Warengruppe", "Attribut")
    Dim col As Long
    For col = 2 To 11
        With wsS.Cells(3, col)
            .Value = headers(col - 2)
            .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 11
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(64, 64, 64)
            .HorizontalAlignment = xlCenter
        End With
    Next col
End Sub

' ================================================================
'  SCHNELLANSICHT - DOPPELKLICK → ARTIKELDETAIL ÖFFNEN
' ================================================================
Sub Schnellansicht_DoppelklickHandler(ByVal Target As Range, Cancel As Boolean)
    If Target.Row < 4 Then Exit Sub
    Cancel = True
    If Target.Worksheet.Cells(Target.Row, SA_COL_ART).Value <> "" Then
        Schnellansicht_Artikelklick Target.Worksheet, Target.Row
    End If
End Sub

' ================================================================
'  SCHNELLANSICHT - ARTIKEL LESEN UND DETAIL ÖFFNEN
' ================================================================
Sub Schnellansicht_Artikelklick(ws As Worksheet, zeile As Long)
    If zeile < 4 Then Exit Sub
    Dim ean   As String : ean   = CStr(ws.Cells(zeile, SA_COL_EAN).Value)
    Dim artNr As String : artNr = CStr(ws.Cells(zeile, SA_COL_ARTNR).Value)
    Dim art   As String : art   = ws.Cells(zeile, SA_COL_ART).Value
    Dim vk    As Double : vk    = Val(ws.Cells(zeile, SA_COL_VK).Value)
    Dim best  As Double : best  = Val(ws.Cells(zeile, SA_COL_BEST).Value)
    Dim einh  As String : einh  = ws.Cells(zeile, SA_COL_EINH).Value
    Dim lag   As String : lag   = ws.Cells(zeile, SA_COL_LAG).Value
    Dim wg    As String : wg    = ws.Cells(zeile, SA_COL_WG).Value
    Dim attr  As String : attr  = ws.Cells(zeile, SA_COL_ATTR).Value
    ' EK aus Artikel-Sheet holen
    Dim ek As Double : ek = 0
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If Not wsA Is Nothing Then
        Dim cEAN2 As Long : cEAN2 = Spalte_Finden(wsA, "EAN13")
        Dim cEK2  As Long : cEK2  = Spalte_Finden(wsA, "EK-PREIS")
        If cEAN2 > 0 And cEK2 > 0 Then
            Dim j As Long
            For j = ART_DATA_START To wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row
                If CStr(wsA.Cells(j, cEAN2).Value) = ean Then
                    ek = Val(wsA.Cells(j, cEK2).Value) : Exit For
                End If
            Next j
        End If
    End If
    g_ReturnSheet = "Schnell"
    ArtikelDetail_Zeigen art, artNr, ean, vk, ek, best, einh, lag, wg, attr, zeile
End Sub

' ================================================================
'  SCHNELLANSICHT - BESTAND DIREKT EDITIEREN (Worksheet_Change)
' ================================================================
Sub Schnellansicht_BestandChange(ByVal Target As Range)
    ' Suchfeld C2: Enter loest Suche aus
    If Target.Row = 2 And Target.Column = 3 Then
        Schnellansicht_Suchen : Exit Sub
    End If
    If Target.Row < 4 Or Target.Column <> SA_COL_BEST Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    Dim wsS As Worksheet : Set wsS = Target.Worksheet
    Dim ean As String : ean = CStr(wsS.Cells(Target.Row, SA_COL_EAN).Value)
    If ean = "" Then Exit Sub
    Dim neuerBestand As Double : neuerBestand = Val(CStr(Target.Value))
    ' Zurückschreiben ins Artikel-Sheet
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Dim cEAN As Long : cEAN = Spalte_Finden(wsA, "EAN13")
    Dim cAnz As Long : cAnz = Spalte_Finden(wsA, "ANZAHL")
    If cEAN = 0 Or cAnz = 0 Then Exit Sub
    Application.EnableEvents = False
    Dim k As Long
    For k = ART_DATA_START To wsA.Cells(wsA.Rows.Count, cEAN).End(xlUp).Row
        If CStr(wsA.Cells(k, cEAN).Value) = ean Then
            wsA.Cells(k, cAnz).Value = neuerBestand
            ' Zeilenfarbe aktualisieren
            If neuerBestand = 0 Then
                wsS.Rows(Target.Row).Interior.Color = RGB(255, 199, 206)
            ElseIf neuerBestand <= 5 Then
                wsS.Rows(Target.Row).Interior.Color = RGB(255, 235, 156)
            Else
                wsS.Rows(Target.Row).Interior.ColorIndex = xlNone
            End If
            Exit For
        End If
    Next k
    Application.EnableEvents = True
End Sub

' ================================================================
'  ARTIKELDETAIL - BLATT BEFÜLLEN UND ANZEIGEN
'  Layout (17 Zeilen):
'   Z1  = Titel
'   Z3  = Artikelname (C3:G3)
'   Z4  = Art.-Nr. (C4, grau) + EAN (F4:G4, grau)
'   Z5  = VK-Preis (C5, gelb) + VK2 (F5, gelb)
'   Z6  = EK-Preis (C6, gelb, versteckt) + EK-Toggle (G6)
'   Z7  = MwSt % (C7, gelb) + Lieferant (F7, gelb)
'   Z8  = Bestand (C8, gelb) + Einheit (F8, gelb)
'   Z9  = Lagerort (C9, gelb) + Warengruppe (F9, gelb)
'   Z10 = Attribut (C10:G10, gelb)
'   Z11 = TextA (C11:G11, gelb)
'   Z12 = TextB (C12:G12, gelb)
'   Z13 = Rohgewinn (auto, grau)
'   Z14 = Trennlinie
'   Z15 = Buttons: SCHLIESSEN (B15), BEWEGUNGEN (D15), SPEICHERN (F15)
'   Z17 = Quellzeile (A17, versteckt)
' ================================================================
Sub ArtikelDetail_Zeigen(art As String, artNr As String, ean As String, _
                         vk As Double, ek As Double, best As Double, _
                         einh As String, lag As String, wg As String, _
                         attr As String, quellZeile As Long)
    Dim wsD As Worksheet : Set wsD = GetSheet("ArtikelDetail")
    If wsD Is Nothing Then
        MsgBox "ArtikelDetail-Blatt fehlt. Bitte Setup_ArtikelDetail ausfuehren.", vbExclamation
        Exit Sub
    End If
    Application.EnableEvents = False

    ' Grundfelder
    wsD.Cells(3, 3).Value  = art
    wsD.Cells(4, 3).Value  = artNr
    wsD.Cells(4, 6).NumberFormat = "@"
    wsD.Cells(4, 6).Value  = CStr(ean)
    wsD.Cells(5, 3).Value  = vk
    wsD.Cells(8, 3).Value  = best
    wsD.Cells(8, 6).Value  = einh
    wsD.Cells(9, 3).Value  = lag
    wsD.Cells(9, 6).Value  = wg
    wsD.Cells(10, 3).Value = attr

    ' EK versteckt anzeigen (Schriftfarbe = Hintergrundfarbe)
    wsD.Cells(6, 3).Value      = ek
    wsD.Cells(6, 3).Font.Color = RGB(255, 255, 200)   ' unsichtbar
    wsD.Cells(6, 7).Value      = "O"                   ' O = EK verborgen
    wsD.Cells(6, 7).Interior.Color = RGB(0, 120, 0)   ' grün

    ' Extra-Felder aus Artikel-Sheet holen (VK2, MwSt, Lieferant, TextA, TextB)
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If Not wsA Is Nothing Then
        Dim cEAN_  As Long : cEAN_    = Spalte_Finden(wsA, "EAN13")
        Dim cVK2   As Long : cVK2     = Spalte_Finden(wsA, "VK2")
        Dim cMwSt  As Long : cMwSt    = Spalte_Finden(wsA, "MwSt")
        Dim cLief  As Long : cLief    = Spalte_Finden(wsA, "Lieferant")
        Dim cTextA As Long : cTextA   = Spalte_Finden(wsA, "TextA")
        Dim cTextB As Long : cTextB   = Spalte_Finden(wsA, "TextB")
        If cEAN_ > 0 Then
            Dim rr As Long
            For rr = ART_DATA_START To wsA.Cells(wsA.Rows.Count, cEAN_).End(xlUp).Row
                If CStr(wsA.Cells(rr, cEAN_).Value) = ean Then
                    If cVK2   > 0 Then wsD.Cells(5, 6).Value  = wsA.Cells(rr, cVK2).Value
                    If cMwSt  > 0 Then wsD.Cells(7, 3).Value  = wsA.Cells(rr, cMwSt).Value
                    If cLief  > 0 Then wsD.Cells(7, 6).Value  = wsA.Cells(rr, cLief).Value
                    If cTextA > 0 Then wsD.Cells(11, 3).Value = wsA.Cells(rr, cTextA).Value
                    If cTextB > 0 Then wsD.Cells(12, 3).Value = wsA.Cells(rr, cTextB).Value
                    Exit For
                End If
            Next rr
        End If
    End If

    ' Rohgewinn live berechnen (VK - EK)
    Dim rohg As Double
    If ek > 0 Then
        rohg = vk - ek
        wsD.Cells(13, 3).Value = Format(rohg, "0.00") & " EUR  (" & _
                                  Format(IIf(vk > 0, rohg / vk * 100, 0), "0.0") & " %)"
    Else
        wsD.Cells(13, 3).Value = "k.A."
    End If

    wsD.Cells(17, 1).Value = quellZeile    ' A17: Quellzeile (versteckt)

    ' Dropdowns befüllen
    ArtikelDetail_SetzeDropdowns wsD
    Application.EnableEvents = True
    wsD.Visible = xlSheetVisible
    wsD.Activate
    wsD.Cells(5, 3).Select
End Sub

' ================================================================
'  ARTIKELDETAIL - DROPDOWNS BEFÜLLEN
' ================================================================
Private Sub ArtikelDetail_SetzeDropdowns(wsD As Worksheet)
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Dim cLag  As Long : cLag  = Spalte_Finden(wsA, "LAGERORT")
    Dim cWG   As Long : cWG   = Spalte_Finden(wsA, "WARENGRUPPE")
    Dim cAttr As Long : cAttr = Spalte_Finden(wsA, "ATTRIBUT")
    Dim cEinh As Long : cEinh = Spalte_Finden(wsA, "EINHEIT")
    Dim cLief As Long : cLief = Spalte_Finden(wsA, "Lieferant")
    Dim lastRow As Long : lastRow = wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row
    Dim dLag  As Object : Set dLag  = CreateObject("Scripting.Dictionary")
    Dim dWG   As Object : Set dWG   = CreateObject("Scripting.Dictionary")
    Dim dAttr As Object : Set dAttr = CreateObject("Scripting.Dictionary")
    Dim dEinh As Object : Set dEinh = CreateObject("Scripting.Dictionary")
    Dim dLief As Object : Set dLief = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim lv As String, wv As String, av As String, ev As String, li As String
    For i = ART_DATA_START To lastRow
        If cLag  > 0 Then lv = Trim(CStr(wsA.Cells(i, cLag).Value))  : If lv  <> "" Then dLag(lv)   = 1
        If cWG   > 0 Then wv = Trim(CStr(wsA.Cells(i, cWG).Value))   : If wv  <> "" Then dWG(wv)    = 1
        If cAttr > 0 Then av = Trim(CStr(wsA.Cells(i, cAttr).Value)) : If av  <> "" Then dAttr(av)  = 1
        If cEinh > 0 Then ev = Trim(CStr(wsA.Cells(i, cEinh).Value)) : If ev  <> "" Then dEinh(ev)  = 1
        If cLief > 0 Then li = Trim(CStr(wsA.Cells(i, cLief).Value)) : If li  <> "" Then dLief(li)  = 1
    Next i
    ' Einheit → F8
    If dEinh.Count > 0 Then
        With wsD.Cells(8, 6).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=Join(dEinh.Keys, ",")
            .ShowError = False
        End With
    End If
    ' Lagerort → C9
    If dLag.Count > 0 Then
        With wsD.Cells(9, 3).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=Join(dLag.Keys, ",")
            .ShowError = False
        End With
    End If
    ' Warengruppe → F9
    If dWG.Count > 0 Then
        With wsD.Cells(9, 6).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=Join(dWG.Keys, ",")
            .ShowError = False
        End With
    End If
    ' Attribut → C10
    If dAttr.Count > 0 Then
        With wsD.Cells(10, 3).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=Join(dAttr.Keys, ",")
            .ShowError = False
        End With
    End If
    ' Lieferant → F7
    If dLief.Count > 0 Then
        With wsD.Cells(7, 6).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=Join(dLief.Keys, ",")
            .ShowError = False
        End With
    End If
    ' MwSt → C7 (feste Liste)
    With wsD.Cells(7, 3).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="0,7,19"
        .ShowError = False
    End With
End Sub

' ================================================================
'  ARTIKELDETAIL - HANDLER (Klick auf Buttons + Cursor)
'  Neues Layout: EK-Toggle = G6, Buttons = Z15
' ================================================================
Sub ArtikelDetail_Handler(ByVal Target As Range)
    ' EK-Toggle Button (G6 = Zeile 6, Spalte 7)
    If Target.Row = 6 And Target.Column = 7 Then
        ArtikelDetail_EK_Toggle Target.Worksheet : Exit Sub
    End If
    ' Cursor sichtbar in editierbaren Feldern (gelbe Zellen)
    Select Case Target.Row
        Case 5  ' VK-Preis (C5) + VK2 (F5)
            If Target.Column = 3 Or Target.Column = 6 Then Application.SendKeys "{F2}" : Exit Sub
        Case 6  ' EK-Preis (C6)
            If Target.Column = 3 Then Application.SendKeys "{F2}" : Exit Sub
        Case 7  ' MwSt (C7) + Lieferant (F7)
            If Target.Column = 3 Or Target.Column = 6 Then Application.SendKeys "{F2}" : Exit Sub
        Case 8  ' Bestand (C8) + Einheit (F8)
            If Target.Column = 3 Or Target.Column = 6 Then Application.SendKeys "{F2}" : Exit Sub
        Case 9  ' Lagerort (C9) + Warengruppe (F9)
            If Target.Column = 3 Or Target.Column = 6 Then Application.SendKeys "{F2}" : Exit Sub
        Case 10 ' Attribut (C10)
            If Target.Column = 3 Then Application.SendKeys "{F2}" : Exit Sub
        Case 11 ' TextA (C11)
            If Target.Column = 3 Then Application.SendKeys "{F2}" : Exit Sub
        Case 12 ' TextB (C12)
            If Target.Column = 3 Then Application.SendKeys "{F2}" : Exit Sub
    End Select
    ' Zeile 15: Buttons (text-basiert)
    If Target.Row = 15 Then
        Dim btn As String : btn = UCase(Trim(CStr(Target.Cells(1, 1).Value)))
        Select Case btn
            Case "SCHLIESSEN"           : ArtikelDetail_Schliessen
            Case "SPEICHERN"            : ArtikelDetail_Speichern
            Case "BEWEGUNGEN", "BEWEGUNGSHISTORIE" : Bewegungen_Zeigen
        End Select
    End If
End Sub

' ================================================================
'  ARTIKELDETAIL - EK-PREIS AUSBLENDEN / EINBLENDEN
' ================================================================
Sub ArtikelDetail_EK_Toggle(wsD As Worksheet)
    Dim ekZelle As Range : Set ekZelle = wsD.Cells(6, 3)   ' EK in C6
    Dim xBtn    As Range : Set xBtn    = wsD.Cells(6, 7)   ' Toggle-Button G6
    If ekZelle.Font.Color = RGB(255, 255, 200) Then
        ' Einblenden
        ekZelle.Font.Color = RGB(80, 80, 80)
        xBtn.Value = "X"
        xBtn.Interior.Color = RGB(180, 0, 0)
    Else
        ' Ausblenden (Schriftfarbe = Hintergrundfarbe = unsichtbar)
        ekZelle.Font.Color = RGB(255, 255, 200)
        xBtn.Value = "O"
        xBtn.Interior.Color = RGB(0, 120, 0)
    End If
End Sub

' ================================================================
'  ARTIKELDETAIL - SCHLIESSEN
'  Springt je nach g_ReturnSheet zurück zur richtigen Seite
' ================================================================
Sub ArtikelDetail_Schliessen()
    Dim wsD As Worksheet : Set wsD = GetSheet("ArtikelDetail")
    If Not wsD Is Nothing Then wsD.Visible = xlSheetHidden
    Dim wsZiel As Worksheet
    If g_ReturnSheet = "Artikel" Then
        Set wsZiel = GetSheet("Artikel")
    Else
        Set wsZiel = GetSheet("Schnell")
    End If
    If Not wsZiel Is Nothing Then wsZiel.Activate
End Sub

' ================================================================
'  ARTIKELDETAIL - SPEICHERN → zurück ins Artikel-Sheet
'  Liest alle Felder aus dem neuen 17-Zeilen-Layout
' ================================================================
Sub ArtikelDetail_Speichern()
    Dim wsD As Worksheet : Set wsD = GetSheet("ArtikelDetail")
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsD Is Nothing Or wsA Is Nothing Then Exit Sub
    On Error GoTo SpErr

    Dim ean As String : ean = CStr(wsD.Cells(4, 6).Value)   ' F4: EAN
    If ean = "" Then MsgBox "Keine EAN - kann nicht speichern.", vbExclamation : Exit Sub

    ' Felder aus neuem Layout lesen
    Dim vkNeu   As Double : vkNeu   = Val(wsD.Cells(5, 3).Value)  ' C5: VK-Preis
    Dim vk2Neu  As Double : vk2Neu  = Val(wsD.Cells(5, 6).Value)  ' F5: VK2
    Dim ekNeu   As Double : ekNeu   = Val(wsD.Cells(6, 3).Value)  ' C6: EK-Preis
    Dim mwstNeu As Double : mwstNeu = Val(wsD.Cells(7, 3).Value)  ' C7: MwSt %
    Dim liefNeu As String : liefNeu = Trim(wsD.Cells(7, 6).Value) ' F7: Lieferant
    Dim bestNeu As Double : bestNeu = Val(wsD.Cells(8, 3).Value)  ' C8: Bestand
    Dim einhNeu As String : einhNeu = Trim(wsD.Cells(8, 6).Value) ' F8: Einheit
    Dim lagNeu  As String : lagNeu  = Trim(wsD.Cells(9, 3).Value) ' C9: Lagerort
    Dim wgNeu   As String : wgNeu   = Trim(wsD.Cells(9, 6).Value) ' F9: Warengruppe
    Dim attrNeu As String : attrNeu = Trim(wsD.Cells(10, 3).Value)' C10: Attribut
    Dim textANeu As String : textANeu = Trim(wsD.Cells(11, 3).Value) ' C11: TextA
    Dim textBNeu As String : textBNeu = Trim(wsD.Cells(12, 3).Value) ' C12: TextB
    Dim qZeile  As Long   : qZeile  = Val(wsD.Cells(17, 1).Value) ' A17: Quellzeile

    Dim cEAN  As Long : cEAN  = Spalte_Finden(wsA, "EAN13")
    Dim cVK   As Long : cVK   = Spalte_Finden(wsA, "VK-PREIS")
    Dim cVK2  As Long : cVK2  = Spalte_Finden(wsA, "VK2")
    Dim cEK   As Long : cEK   = Spalte_Finden(wsA, "EK-PREIS")
    Dim cMwSt As Long : cMwSt = Spalte_Finden(wsA, "MwSt")
    Dim cLief As Long : cLief = Spalte_Finden(wsA, "Lieferant")
    Dim cAnz  As Long : cAnz  = Spalte_Finden(wsA, "ANZAHL")
    Dim cEinh As Long : cEinh = Spalte_Finden(wsA, "EINHEIT")
    Dim cLag  As Long : cLag  = Spalte_Finden(wsA, "LAGERORT")
    Dim cWG   As Long : cWG   = Spalte_Finden(wsA, "WARENGRUPPE")
    Dim cAttr As Long : cAttr = Spalte_Finden(wsA, "ATTRIBUT")
    Dim cTA   As Long : cTA   = Spalte_Finden(wsA, "TextA")
    Dim cTB   As Long : cTB   = Spalte_Finden(wsA, "TextB")
    If cEAN = 0 Then MsgBox "EAN-Spalte nicht gefunden.", vbExclamation : Exit Sub

    Application.EnableEvents = False
    Dim k As Long
    For k = ART_DATA_START To wsA.Cells(wsA.Rows.Count, cEAN).End(xlUp).Row
        If CStr(wsA.Cells(k, cEAN).Value) = ean Then
            If cVK   > 0 Then wsA.Cells(k, cVK).Value   = vkNeu
            If cVK2  > 0 Then wsA.Cells(k, cVK2).Value  = vk2Neu
            If cEK   > 0 Then wsA.Cells(k, cEK).Value   = ekNeu
            If cMwSt > 0 Then wsA.Cells(k, cMwSt).Value = mwstNeu
            If cLief > 0 And liefNeu  <> "" Then wsA.Cells(k, cLief).Value  = liefNeu
            If cAnz  > 0 Then wsA.Cells(k, cAnz).Value  = bestNeu
            If cEinh > 0 And einhNeu  <> "" Then wsA.Cells(k, cEinh).Value  = einhNeu
            If cLag  > 0 And lagNeu   <> "" Then wsA.Cells(k, cLag).Value   = lagNeu
            If cWG   > 0 And wgNeu    <> "" Then wsA.Cells(k, cWG).Value    = wgNeu
            If cAttr > 0 Then wsA.Cells(k, cAttr).Value = attrNeu
            If cTA   > 0 Then wsA.Cells(k, cTA).Value   = textANeu
            If cTB   > 0 Then wsA.Cells(k, cTB).Value   = textBNeu
            Exit For
        End If
    Next k

    ' Schnellansicht-Zeile mitaktualisieren
    If g_ReturnSheet = "Schnell" Then
        Dim wsSc As Worksheet : Set wsSc = GetSheet("Schnell")
        If Not wsSc Is Nothing And qZeile >= 4 Then
            wsSc.Cells(qZeile, SA_COL_VK).Value   = vkNeu
            wsSc.Cells(qZeile, SA_COL_BEST).Value  = bestNeu
            wsSc.Cells(qZeile, SA_COL_EINH).Value  = einhNeu
            wsSc.Cells(qZeile, SA_COL_LAG).Value   = lagNeu
            wsSc.Cells(qZeile, SA_COL_WG).Value    = wgNeu
            wsSc.Cells(qZeile, SA_COL_ATTR).Value  = attrNeu
        End If
    End If
    Application.EnableEvents = True

    wsD.Visible = xlSheetHidden
    ArtikelDetail_Schliessen   ' nutzt g_ReturnSheet für Navigation
    Exit Sub
SpErr:
    Application.EnableEvents = True
    MsgBox "Fehler beim Speichern: " & Err.Description, vbExclamation
End Sub

' ================================================================
'  SETUP: Schnellansicht neu aufbauen
' ================================================================
Sub Setup_Schnellansicht()
    Dim wsS As Worksheet : Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then
        MsgBox "Schnellansicht-Blatt nicht gefunden!" & Chr(10) & SheetListe(), vbCritical : Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Spaltenbreiten
    wsS.Columns("A").ColumnWidth = 3
    wsS.Columns("B").ColumnWidth = 5
    wsS.Columns("C").ColumnWidth = 12
    wsS.Columns("D").ColumnWidth = 30
    wsS.Columns("E").ColumnWidth = 16
    wsS.Columns("F").ColumnWidth = 10
    wsS.Columns("G").ColumnWidth = 10
    wsS.Columns("H").ColumnWidth = 8
    wsS.Columns("I").ColumnWidth = 14
    wsS.Columns("J").ColumnWidth = 14
    wsS.Columns("K").ColumnWidth = 14

    ' Zeilenhöhen
    wsS.Rows(1).RowHeight = 28
    wsS.Rows(2).RowHeight = 24
    wsS.Rows(3).RowHeight = 22

    ' Zeile 1: Titel
    On Error Resume Next : wsS.Range("A1:K1").UnMerge : On Error GoTo 0
    wsS.Range("A1:K1").Merge
    With wsS.Cells(1, 1)
        .Value = "Schnellansicht - Artikelbestand"
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Arial" : .Font.Size = 14 : .Font.Bold = True
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With

    ' Zeile 2: Suchfeld + Buttons
    wsS.Cells(2, 1).Value = ""
    wsS.Cells(2, 2).Value = "Suche:"
    wsS.Cells(2, 2).Font.Bold = True : wsS.Cells(2, 2).Font.Name = "Arial"
    On Error Resume Next : wsS.Range("C2:G2").UnMerge : On Error GoTo 0
    wsS.Range("C2:G2").Merge
    With wsS.Cells(2, 3)
        .Interior.Color = RGB(255, 255, 200)
        .Font.Name = "Arial" : .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .Borders(xlEdgeBottom).LineStyle = xlContinuous : .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).LineStyle = xlContinuous    : .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).LineStyle = xlContinuous   : .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous  : .Borders(xlEdgeRight).Weight = xlMedium
    End With
    ' H2: SUCHEN
    With wsS.Cells(2, 8)
        .Value = "SUCHEN" : .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 11
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(0, 120, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With
    ' I2: LEEREN
    With wsS.Cells(2, 9)
        .Value = "LEEREN" : .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 11
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(180, 90, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With
    ' J2: Treffer-Anzeige
    With wsS.Cells(2, 10)
        .Font.Name = "Arial" : .Font.Size = 11 : .Font.Bold = True
        .HorizontalAlignment = xlCenter : .Interior.ColorIndex = xlNone
    End With
    ' K2: AKTUALISIEREN
    With wsS.Cells(2, 11)
        .Value = "AKTUALISIEREN" : .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 10
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(64, 64, 64)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' Autofilter entfernen falls vorhanden
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False

    ' Zeile 3: Spaltenköpfe
    Schnellansicht_Headers wsS

    ' Zeilen 1-3 einfrieren
    wsS.Activate
    wsS.Cells(4, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ' Grundformat
    With wsS.Cells
        .Font.Name = "Arial" : .Font.Size = 11
    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Schnellansicht-Layout aktualisiert!" & Chr(10) & Chr(10) & _
           "Sheet-Modul-Code bitte einfuegen (Alt+F11 → Schnell):" & Chr(10) & Chr(10) & _
           "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10) & _
           "    Schnellansicht_Handler Target" & Chr(10) & "End Sub" & Chr(10) & Chr(10) & _
           "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10) & _
           "    Schnellansicht_DoppelklickHandler Target, Cancel" & Chr(10) & "End Sub" & Chr(10) & Chr(10) & _
           "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10) & _
           "    Schnellansicht_BestandChange Target" & Chr(10) & "End Sub", _
           vbInformation, "Setup_Schnellansicht"
End Sub

' ================================================================
'  SETUP: ArtikelDetail-Blatt erstellen / neu formatieren (17-Zeilen-Layout)
' ================================================================
Sub Setup_ArtikelDetail()
    Dim wsD As Worksheet : Set wsD = GetSheet("ArtikelDetail")
    If wsD Is Nothing Then
        Set wsD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsD.Name = "ArtikelDetail"
        MsgBox "Neues ArtikelDetail-Blatt angelegt." & Chr(10) & _
               "Bitte Sheet-Modul-Code einfuegen (Alt+F11 -> ArtikelDetail).", _
               vbInformation, "Hinweis"
    End If
    Application.DisplayAlerts = False
    wsD.Visible = xlSheetVisible
    wsD.Cells.UnMerge : wsD.Cells.ClearContents : wsD.Cells.ClearFormats
    wsD.Cells.Validation.Delete
    Application.DisplayAlerts = True

    wsD.Cells.Font.Name = "Arial" : wsD.Cells.Font.Size = 12 : wsD.Cells.RowHeight = 22
    wsD.Columns("A").ColumnWidth = 3   ' leer (Rand)
    wsD.Columns("B").ColumnWidth = 14  ' Label
    wsD.Columns("C").ColumnWidth = 22  ' Wert links
    wsD.Columns("D").ColumnWidth = 2   ' Abstand
    wsD.Columns("E").ColumnWidth = 14  ' Label rechts
    wsD.Columns("F").ColumnWidth = 22  ' Wert rechts
    wsD.Columns("G").ColumnWidth = 5   ' Toggle-Button

    ' ── Zeile 1: Titel ─────────────────────────────────────────────
    wsD.Rows(1).RowHeight = 30
    wsD.Range("A1:G1").Merge
    With wsD.Cells(1, 1)
        .Value = "Artikel - Details"
        .Interior.Color = RGB(64, 64, 64) : .Font.Color = RGB(255, 255, 255)
        .Font.Size = 14 : .Font.Bold = True
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With

    ' ── Zeile 3: Artikelname ────────────────────────────────────────
    wsD.Rows(3).RowHeight = 40
    wsD.Cells(3, 2).Value = "Artikel:" : wsD.Cells(3, 2).Font.Bold = True
    wsD.Range("C3:G3").Merge
    With wsD.Cells(3, 3)
        .HorizontalAlignment = xlLeft : .VerticalAlignment = xlTop : .WrapText = True
    End With

    ' ── Zeile 4: Art.-Nr. + EAN (nur Anzeige, grau) ────────────────
    wsD.Cells(4, 2).Value = "Art.-Nr.:" : wsD.Cells(4, 2).Font.Bold = True
    With wsD.Cells(4, 3) : .Interior.Color = RGB(240, 240, 240) : End With
    wsD.Cells(4, 5).Value = "EAN:" : wsD.Cells(4, 5).Font.Bold = True
    wsD.Range("F4:G4").Merge
    With wsD.Cells(4, 6) : .Interior.Color = RGB(240, 240, 240) : .NumberFormat = "@" : End With

    ' ── Zeile 5: VK-Preis + VK2 (editierbar, gelb) ─────────────────
    wsD.Cells(5, 2).Value = "VK-Preis:" : wsD.Cells(5, 2).Font.Bold = True
    With wsD.Cells(5, 3)
        .Interior.Color = RGB(255, 255, 200) : .NumberFormat = "0.00"
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(5, 5).Value = "VK2 (opt.):" : wsD.Cells(5, 5).Font.Bold = True
    With wsD.Cells(5, 6)
        .Interior.Color = RGB(255, 255, 200) : .NumberFormat = "0.00"
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(5, 7).Value = ""

    ' ── Zeile 6: EK-Preis (editierbar, versteckt) + Toggle-Button ──
    wsD.Cells(6, 2).Value = "EK-Preis:" : wsD.Cells(6, 2).Font.Bold = True
    With wsD.Cells(6, 3)
        .Interior.Color = RGB(255, 255, 200) : .NumberFormat = "0.00"
        .Font.Color = RGB(255, 255, 200)    ' startet UNSICHTBAR (Schrift=Hintergrund)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(6, 5).Value = ""
    wsD.Range("F6:F6").Interior.ColorIndex = xlNone
    ' G6: EK-Toggle-Button → startet als "O" (grün = EK verborgen)
    With wsD.Cells(6, 7)
        .Value = "O"
        .Font.Bold = True : .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 120, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' ── Zeile 7: MwSt % + Lieferant (editierbar) ───────────────────
    wsD.Cells(7, 2).Value = "MwSt %:" : wsD.Cells(7, 2).Font.Bold = True
    With wsD.Cells(7, 3)
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(7, 5).Value = "Lieferant:" : wsD.Cells(7, 5).Font.Bold = True
    With wsD.Cells(7, 6)
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(7, 7).Value = ""

    ' ── Zeile 8: Bestand + Einheit (editierbar) ────────────────────
    wsD.Cells(8, 2).Value = "Bestand:" : wsD.Cells(8, 2).Font.Bold = True
    With wsD.Cells(8, 3)
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(8, 5).Value = "Einheit:" : wsD.Cells(8, 5).Font.Bold = True
    With wsD.Cells(8, 6)
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(8, 7).Value = ""

    ' ── Zeile 9: Lagerort + Warengruppe (editierbar) ───────────────
    wsD.Cells(9, 2).Value = "Lagerort:" : wsD.Cells(9, 2).Font.Bold = True
    With wsD.Cells(9, 3)
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(9, 5).Value = "Warengruppe:" : wsD.Cells(9, 5).Font.Bold = True
    With wsD.Cells(9, 6)
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Cells(9, 7).Value = ""

    ' ── Zeile 10: Attribut (editierbar) ────────────────────────────
    wsD.Cells(10, 2).Value = "Attribut:" : wsD.Cells(10, 2).Font.Bold = True
    wsD.Range("C10:G10").Merge
    With wsD.Cells(10, 3)
        .Interior.Color = RGB(255, 255, 200)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With

    ' ── Zeile 11: TextA (editierbar) ───────────────────────────────
    wsD.Cells(11, 2).Value = "TextA:" : wsD.Cells(11, 2).Font.Bold = True
    wsD.Range("C11:G11").Merge
    With wsD.Cells(11, 3)
        .Interior.Color = RGB(255, 255, 200)
        .WrapText = True
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Rows(11).RowHeight = 32

    ' ── Zeile 12: TextB (editierbar) ───────────────────────────────
    wsD.Cells(12, 2).Value = "TextB:" : wsD.Cells(12, 2).Font.Bold = True
    wsD.Range("C12:G12").Merge
    With wsD.Cells(12, 3)
        .Interior.Color = RGB(255, 255, 200)
        .WrapText = True
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlThin
    End With
    wsD.Rows(12).RowHeight = 32

    ' ── Zeile 13: Rohgewinn (auto, grau) ───────────────────────────
    wsD.Cells(13, 2).Value = "Rohgewinn:" : wsD.Cells(13, 2).Font.Bold = True
    wsD.Range("C13:G13").Merge
    With wsD.Cells(13, 3)
        .Interior.Color = RGB(240, 240, 240)
        .Font.Color = RGB(80, 80, 80)
        .HorizontalAlignment = xlLeft
    End With

    ' ── Zeile 14: Trennlinie ───────────────────────────────────────
    wsD.Rows(14).RowHeight = 4
    wsD.Range("A14:G14").Interior.Color = RGB(180, 180, 180)

    ' ── Zeile 15: Buttons ──────────────────────────────────────────
    wsD.Rows(15).RowHeight = 28
    With wsD.Cells(15, 2)                    ' B15: SCHLIESSEN
        .Value = "SCHLIESSEN" : .Font.Bold = True : .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With
    With wsD.Cells(15, 4)                    ' D15: BEWEGUNGEN
        .Value = "BEWEGUNGEN" : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 80, 160)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With
    With wsD.Cells(15, 6)                    ' F15: SPEICHERN
        .Value = "SPEICHERN" : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 120, 0)
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' ── Zeile 17: Quellzeile (versteckt in A17) ────────────────────
    wsD.Rows(17).RowHeight = 4
    wsD.Cells(17, 1).Font.Color = RGB(255, 255, 255)

    wsD.Visible = xlSheetHidden
    MsgBox "ArtikelDetail-Blatt neu aufgebaut (17-Zeilen-Layout)!" & Chr(10) & Chr(10) & _
           "Sheet-Modul-Code einfuegen (Alt+F11 → ArtikelDetail):" & Chr(10) & Chr(10) & _
           "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10) & _
           "    ArtikelDetail_Handler Target" & Chr(10) & "End Sub" & Chr(10) & Chr(10) & _
           "Strg+S speichern.", vbInformation, "Setup_ArtikelDetail"
End Sub

' ================================================================
'  ARTIKELBLATT - FARBCODIERUNG LAGERBESTAND ENTFERNEN
'  Entfernt alle roten/gelben Bestandsfarben aus dem Artikelblatt
' ================================================================
Sub Artikelblatt_FarbcodeEntfernen()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Dim lastRow As Long : lastRow = wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    Dim c As Long
    For i = ART_DATA_START To lastRow
        c = wsA.Rows(i).Interior.Color
        ' Nur Bestandsfarben entfernen (Rot und Gelb), nicht die gelbe Markierungszeile
        If c = RGB(255, 199, 206) Or c = RGB(255, 235, 156) Or _
           c = RGB(255, 0, 0)     Or c = RGB(255, 255, 0) Then
            wsA.Rows(i).Interior.ColorIndex = xlNone
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "Farbcodierung aus Artikelblatt entfernt.", vbInformation, "Fertig"
End Sub

' ================================================================
'  ARTIKELBLATT - SUCHE
'  Filtert Artikelblatt-Zeilen anhand des Suchfelds in Zeile 1
'  Suchfeld: Zellen B1:F1 (merged), SUCHEN-Button daneben
' ================================================================
Sub Artikelblatt_Suche()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Dim such As String : such = Trim(wsA.Cells(2, 2).Value)   ' B2: Suchfeld
    If such = "" Then Artikelblatt_FilterLoeschen : Exit Sub

    Dim woerter()  As String : woerter   = Split(LCase(such), " ")
    Dim nurZahlen  As Boolean : nurZahlen = (such = CStr(Val(such)) And Val(such) > 0)

    Dim cArt  As Long : cArt  = Spalte_Finden(wsA, "ARTIKEL")
    Dim cNr   As Long : cNr   = Spalte_Finden(wsA, "ARTIKELNR")
    Dim cEAN  As Long : cEAN  = Spalte_Finden(wsA, "EAN13")
    If cArt = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim lastRow As Long : lastRow = wsA.Cells(wsA.Rows.Count, cArt).End(xlUp).Row
    Dim treffer As Long : treffer = 0
    Dim i As Long, w As Integer, passt As Boolean

    For i = ART_DATA_START To lastRow
        Dim suchIn As String
        If nurZahlen Then
            suchIn = LCase(CStr(wsA.Cells(i, cArt).Value) & " " & _
                          CStr(wsA.Cells(i, cNr).Value) & " " & _
                          CStr(wsA.Cells(i, cEAN).Value))
        Else
            suchIn = LCase(CStr(wsA.Cells(i, cArt).Value))
        End If
        passt = True
        For w = 0 To UBound(woerter)
            If Trim(woerter(w)) <> "" Then
                If InStr(suchIn, Trim(woerter(w))) = 0 Then passt = False : Exit For
            End If
        Next w
        wsA.Rows(i).Hidden = Not passt
        If passt Then treffer = treffer + 1
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Zur ersten sichtbaren Trefferzeile scrollen
    Dim firstHit As Long : firstHit = 0
    Dim fi As Long
    For fi = ART_DATA_START To lastRow
        If Not wsA.Rows(fi).Hidden Then
            firstHit = fi
            Exit For
        End If
    Next fi
    If firstHit > 0 Then
        Application.Goto wsA.Cells(firstHit, 2), True
    Else
        MsgBox "Keine Treffer gefunden.", vbInformation, "Suche"
    End If
End Sub

' ================================================================
'  ARTIKELBLATT - FILTER LÖSCHEN
' ================================================================
Sub Artikelblatt_FilterLoeschen()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim lastRow As Long : lastRow = wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    For i = ART_DATA_START To lastRow
        wsA.Rows(i).Hidden = False
    Next i
    wsA.Cells(2, 2).Value = ""   ' Suchfeld B2 leeren
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ' Zurück zur ersten Datenzeile scrollen
    Application.Goto wsA.Cells(ART_DATA_START, 2), True
End Sub

' ================================================================
'  ARTIKELBLATT - SUCHE BEI EINGABE (Worksheet_Change Handler)
'  Wird aus Sheet-Modul aufgerufen wenn sich B1 ändert
' ================================================================
Sub Artikelblatt_SearchChange(ByVal Target As Range)
    If Target.Row = 2 And Target.Column = 2 Then
        Artikelblatt_Suche
    End If
End Sub

' ================================================================
'  ARTIKELBLATT - DOPPELKLICK → ARTIKELDETAIL ÖFFNEN
'  Wird aus Worksheet_BeforeDoubleClick des Artikel-Sheets aufgerufen
' ================================================================
Sub Artikelblatt_DoppelklickHandler(ByVal Target As Range, Cancel As Boolean)
    Cancel = True   ' Excel-Bearbeitungsmodus IMMER verhindern (alle Zeilen)

    ' Nur Datenzeilen (keine Titel/Suche/Button/Header-Zeilen)
    If Target.Row < ART_DATA_START Then Exit Sub

    ' Prüfung über EAN-Spalte: nur öffnen wenn echte EAN vorhanden
    Dim cEAN As Long : cEAN = Spalte_Finden(Target.Worksheet, "EAN13")
    If cEAN = 0 Then cEAN = 2   ' Fallback: Spalte B
    Dim eanVal As String : eanVal = Trim(CStr(Target.Worksheet.Cells(Target.Row, cEAN).Value))

    ' Leere Zeile oder Headertext → kein Artikel
    If eanVal = "" Then Exit Sub
    If UCase(eanVal) = "EAN13" Then Exit Sub   ' Headerzeile abfangen

    Artikelblatt_Artikelklick Target.Worksheet, Target.Row
End Sub

' ================================================================
'  ARTIKELBLATT - ARTIKEL LESEN UND DETAIL ÖFFNEN
' ================================================================
Sub Artikelblatt_Artikelklick(ws As Worksheet, zeile As Long)
    If zeile < ART_DATA_START Then Exit Sub
    Dim cEAN  As Long : cEAN  = Spalte_Finden(ws, "EAN13")
    Dim cArt  As Long : cArt  = Spalte_Finden(ws, "ARTIKEL")
    Dim cNr   As Long : cNr   = Spalte_Finden(ws, "ARTIKELNR")
    Dim cVK   As Long : cVK   = Spalte_Finden(ws, "VK-PREIS")
    Dim cAnz  As Long : cAnz  = Spalte_Finden(ws, "ANZAHL")
    Dim cEinh As Long : cEinh = Spalte_Finden(ws, "EINHEIT")
    Dim cLag  As Long : cLag  = Spalte_Finden(ws, "LAGERORT")
    Dim cWG   As Long : cWG   = Spalte_Finden(ws, "WARENGRUPPE")
    Dim cAttr As Long : cAttr = Spalte_Finden(ws, "ATTRIBUT")
    Dim cEK   As Long : cEK   = Spalte_Finden(ws, "EK-PREIS")

    Dim ean   As String : ean   = CStr(ws.Cells(zeile, cEAN).Value)
    Dim art   As String : art   = ws.Cells(zeile, cArt).Value
    Dim artNr As String : artNr = CStr(ws.Cells(zeile, cNr).Value)
    Dim vk    As Double : vk    = 0
    Dim ek    As Double : ek    = 0
    Dim best  As Double : best  = 0
    On Error Resume Next
    If cVK  > 0 Then vk   = CDbl(ws.Cells(zeile, cVK).Value)
    If cEK  > 0 Then ek   = CDbl(ws.Cells(zeile, cEK).Value)
    If cAnz > 0 Then best = CDbl(ws.Cells(zeile, cAnz).Value)
    On Error GoTo 0
    Dim einh  As String : einh  = IIf(cEinh > 0, CStr(ws.Cells(zeile, cEinh).Value), "")
    Dim lag   As String : lag   = IIf(cLag  > 0, CStr(ws.Cells(zeile, cLag).Value), "")
    Dim wg    As String : wg    = IIf(cWG   > 0, CStr(ws.Cells(zeile, cWG).Value), "")
    Dim attr  As String : attr  = IIf(cAttr > 0, CStr(ws.Cells(zeile, cAttr).Value), "")

    g_ReturnSheet = "Artikel"
    ArtikelDetail_Zeigen art, artNr, ean, vk, ek, best, einh, lag, wg, attr, zeile
End Sub

' ================================================================
'  ARTIKELBLATT - FEHLENDE FELDER ERGÄNZEN
'  Fügt neue Spalten (TextA, TextB, VK2, MwSt, Lieferant, WE-Datum)
'  ans Ende des Artikelblattes, wenn sie noch nicht existieren
' ================================================================
Sub Artikelblatt_FelderErgaenzen()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub

    ' Liste der benötigten Spalten
    Dim felder As Variant
    felder = Array("VK2", "MwSt", "Lieferant", "TextA", "TextB", "WE-Datum")

    Dim f As Integer
    Dim lastCol As Long : lastCol = wsA.Cells(ART_HEADER_ROW, wsA.Columns.Count).End(xlToLeft).Column
    For f = 0 To UBound(felder)
        If Spalte_Finden(wsA, CStr(felder(f))) = 0 Then
            lastCol = lastCol + 1
            wsA.Cells(ART_HEADER_ROW, lastCol).Value = CStr(felder(f))
            wsA.Cells(ART_HEADER_ROW, lastCol).Font.Bold = True
            wsA.Cells(ART_HEADER_ROW, lastCol).Interior.Color = RGB(64, 64, 64)
            wsA.Cells(ART_HEADER_ROW, lastCol).Font.Color = RGB(255, 255, 255)
            wsA.Columns(lastCol).Hidden = True   ' neue Spalten versteckt
        End If
    Next f
    MsgBox "Felder geprüft/ergänzt (neue Spalten ausgeblendet)." & Chr(10) & _
           "Strg+S speichern.", vbInformation, "FelderErgaenzen"
End Sub

' ================================================================
'  SETUP: Artikelblatt-Layout aufbauen
'  Zeile 1 = Suchfeld (B1:F1 merged) + SUCHEN (G1) + LEEREN (H1)
'  Zeile 2 = Buttons (Zu/Abgang, Etikett, EK, Filter, Schnell, Neu)
'  Zeile 3 = Spaltenköpfe
'  Zeile 4+ = Daten
'  Spalte A = immer leer (Rand)
' ================================================================
Sub Setup_Artikelblatt()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then
        MsgBox "Artikel-Sheet nicht gefunden!" & Chr(10) & SheetListe(), vbCritical : Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    ' ── Titelzeile einfügen (nur wenn noch nicht vorhanden) ────────
    Dim hatTitel As Boolean
    hatTitel = (InStr(1, CStr(wsA.Cells(1, 1).Value), "Artikel", vbTextCompare) > 0 _
             Or InStr(1, CStr(wsA.Cells(1, 2).Value), "Artikel", vbTextCompare) > 0)

    If Not hatTitel Then
        wsA.Rows(1).Insert Shift:=xlDown   ' alles eine Zeile tiefer
    End If

    ' ── Zeile 1: Titelzeile ────────────────────────────────────────
    wsA.Rows(1).RowHeight = 30
    On Error Resume Next : wsA.Rows(1).UnMerge : On Error GoTo 0
    wsA.Range("A1:M1").Merge
    With wsA.Cells(1, 1)
        .Value = "Artikel - Lagerverwaltung"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(40, 40, 100)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With

    ' ── Zeile 2: Suchzeile komplett neu aufbauen (alte Reste löschen) ─
    On Error Resume Next : wsA.Rows(2).UnMerge : On Error GoTo 0
    wsA.Rows(2).ClearContents
    wsA.Rows(2).ClearFormats
    wsA.Rows(2).RowHeight = 28
    ' A2: "Suche:" Label
    With wsA.Cells(2, 1)
        .Value = "Suche:"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 10
        .HorizontalAlignment = xlRight : .VerticalAlignment = xlCenter
        .Interior.ColorIndex = xlNone
    End With
    wsA.Columns("A").ColumnWidth = 8
    On Error Resume Next : wsA.Range("B2:F2").UnMerge : On Error GoTo 0
    wsA.Range("B2:F2").Merge
    With wsA.Cells(2, 2)                             ' B2: Suchfeld
        .Interior.Color = RGB(255, 255, 200)
        .Font.Name = "Arial" : .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With
    With wsA.Cells(2, 7)                             ' G2: SUCHEN
        .Value = "SUCHEN"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 120, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With
    With wsA.Cells(2, 8)                             ' H2: LEEREN
        .Value = "LEEREN"
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(180, 90, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With

    ' ── Zeile 3: Button-Zeile komplett leeren + neu aufbauen ────────
    wsA.Rows(3).ClearContents
    wsA.Rows(3).ClearFormats
    wsA.Rows(3).RowHeight = 28
    Dim btnDefs As Variant
    btnDefs = Array( _
        Array(3,  "ZU-/ABGANG",     RGB(64, 64, 64)), _
        Array(5,  "ETIKETT",        RGB(64, 64, 64)), _
        Array(7,  "EK",             RGB(64, 64, 64)), _
        Array(9,  "FILTER LOESCHEN", RGB(200, 80, 0)), _
        Array(11, "SCHNELLANSICHT", RGB(0, 80, 160)), _
        Array(13, "NEUER ARTIKEL",  RGB(0, 100, 0)))
    Dim b As Integer
    For b = 0 To UBound(btnDefs)
        With wsA.Cells(3, btnDefs(b)(0))
            .Value = btnDefs(b)(1)
            .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 10
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = btnDefs(b)(2)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
        End With
    Next b

    ' ── Zeile 4: Spaltenköpfe (ART_HEADER_ROW = 4) ─────────────────
    wsA.Rows(4).RowHeight = 22

    ' ── Spalte A: Nr.-Spalte (Laufnummer) ──────────────────────────
    wsA.Columns("A").ColumnWidth = 5
    wsA.Columns("A").Interior.ColorIndex = xlNone
    With wsA.Cells(ART_HEADER_ROW, 1)
        .Value = "Nr."
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 10
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        .Interior.Color = RGB(30, 60, 120)
        .Font.Color = RGB(255, 255, 255)
    End With
    ' Spalte B (EAN13) sichtbar machen
    wsA.Columns("B").ColumnWidth = 16

    ' ── Alle gelben Zeilen entfernen ───────────────────────────────
    Dim clearR As Long
    For clearR = 2 To wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row
        If wsA.Rows(clearR).Interior.Color = RGB(255, 255, 153) Then
            wsA.Rows(clearR).Interior.ColorIndex = xlNone
        End If
    Next clearR

    ' ── Nr.-Spalte auffüllen ──────────────────────────────────────
    Artikel_NrAuffuellen wsA

    ' ── Zeilen 1-4 einfrieren ──────────────────────────────────────
    wsA.Activate
    wsA.Cells(ART_DATA_START, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    MsgBox "Artikelblatt-Layout aufgebaut!" & Chr(10) & Chr(10) & _
           "Titelzeile (Zeile 1), Suche (Zeile 2), Buttons (Zeile 3)," & Chr(10) & _
           "Kopfzeile (Zeile 4), Daten ab Zeile 5." & Chr(10) & Chr(10) & _
           "Jetzt: Artikelblatt_FelderErgaenzen ausfuehren!" & Chr(10) & _
           "Dann: Setup_ArtikelDetail ausfuehren!" & Chr(10) & _
           "Dann: Strg+S speichern.", vbInformation, "Setup_Artikelblatt"
End Sub

' ================================================================
'  DIAGNOSE: Zeigt welche Spalten in welcher Zeile gefunden werden
' ================================================================
Sub Artikel_Diagnose()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then MsgBox "Artikel-Sheet nicht gefunden!" : Exit Sub

    Dim msg As String
    msg = "=== ARTIKEL-SHEET DIAGNOSE ===" & Chr(10) & Chr(10)

    ' Zeile 1 Inhalt
    msg = msg & "Zeile 1, A1: " & CStr(wsA.Cells(1, 1).Value) & Chr(10)
    msg = msg & "Zeile 1, B1: " & CStr(wsA.Cells(1, 2).Value) & Chr(10)
    msg = msg & "Zeile 3, A3: " & CStr(wsA.Cells(3, 1).Value) & Chr(10)
    msg = msg & "Zeile 3, B3: " & CStr(wsA.Cells(3, 2).Value) & Chr(10)
    msg = msg & "Zeile 4, A4: " & CStr(wsA.Cells(4, 1).Value) & Chr(10)
    msg = msg & "Zeile 4, B4: " & CStr(wsA.Cells(4, 2).Value) & Chr(10)
    msg = msg & "Zeile 4, C4: " & CStr(wsA.Cells(4, 3).Value) & Chr(10)
    msg = msg & "Zeile 5, A5: " & CStr(wsA.Cells(5, 1).Value) & Chr(10)
    msg = msg & "Zeile 5, B5: " & CStr(wsA.Cells(5, 2).Value) & Chr(10)
    msg = msg & Chr(10) & "ART_HEADER_ROW = " & ART_HEADER_ROW & Chr(10)
    msg = msg & "ART_DATA_START = " & ART_DATA_START & Chr(10) & Chr(10)

    ' Spalte_Finden Ergebnisse
    Dim felder As Variant
    felder = Array("EAN13", "ARTIKEL", "ARTIKELNR", "VK-PREIS", "EK-PREIS", "ANZAHL")
    Dim f As Integer
    msg = msg & "Spalten gefunden:" & Chr(10)
    For f = 0 To UBound(felder)
        Dim col As Long : col = Spalte_Finden(wsA, CStr(felder(f)))
        msg = msg & "  " & felder(f) & " → Spalte " & IIf(col = 0, "NICHT GEFUNDEN!", col) & Chr(10)
    Next f

    MsgBox msg, vbInformation, "Diagnose"
End Sub

' ================================================================
'  SPALTE A BEREINIGEN: EAN aus A in richtige Spalte verschieben
' ================================================================
Sub Artikel_SpalteA_Bereinigen()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub

    Dim cEAN As Long : cEAN = Spalte_Finden(wsA, "EAN13")
    If cEAN = 0 Or cEAN = 1 Then
        MsgBox "EAN13-Spalte nicht gefunden oder bereits in Spalte A." & Chr(10) & _
               "Bitte zuerst Artikel_Diagnose ausführen.", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long : lastRow = wsA.Cells(wsA.Rows.Count, cEAN).End(xlUp).Row
    Dim fixAnz As Long : fixAnz = 0
    Application.EnableEvents = False

    Dim i As Long
    For i = ART_DATA_START To lastRow
        ' Wenn A nicht leer UND EAN-Spalte leer → EAN von A in cEAN verschieben
        If wsA.Cells(i, 1).Value <> "" And wsA.Cells(i, cEAN).Value = "" Then
            wsA.Cells(i, cEAN).Value = wsA.Cells(i, 1).Value
            wsA.Cells(i, 1).ClearContents
            fixAnz = fixAnz + 1
        ElseIf wsA.Cells(i, 1).Value <> "" Then
            ' A hat Inhalt aber EAN-Spalte auch schon → A einfach leeren
            wsA.Cells(i, 1).ClearContents
            fixAnz = fixAnz + 1
        End If
    Next i

    Application.EnableEvents = True
    MsgBox fixAnz & " Zeilen bereinigt. Spalte A ist jetzt leer." & Chr(10) & _
           "Strg+S speichern.", vbInformation, "Spalte A Bereinigung"
End Sub


' ================================================================
'  NR.-SPALTE AUFFÜLLEN  (Spalte A mit 1, 2, 3 ... befüllen)
'  Kann auch separat als Makro ausgeführt werden
' ================================================================
Sub Artikel_AlleZeigenFix()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then MsgBox "Artikel-Sheet nicht gefunden!" : Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wsA.Activate
    ' Alle Zeilen 5 bis 10000 zwangsweise einblenden
    wsA.Rows("5:10000").Hidden = False
    wsA.Cells(2, 2).Value = ""   ' Suchfeld leeren
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Goto wsA.Cells(5, 2), True
    MsgBox "Fertig – alle Artikel eingeblendet.", vbInformation, "OK"
End Sub

Sub Artikel_NrAuffuellen(Optional wsIn As Worksheet = Nothing)
    Dim wsA As Worksheet
    If wsIn Is Nothing Then
        Set wsA = GetSheet("Artikel")
    Else
        Set wsA = wsIn
    End If
    If wsA Is Nothing Then Exit Sub

    Dim cArt As Long : cArt = Spalte_Finden(wsA, "ARTIKEL")
    If cArt = 0 Then cArt = 3   ' Fallback

    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    ' Header "Nr." in Zeile 4, Spalte A
    With wsA.Cells(ART_HEADER_ROW, 1)
        .Value = "Nr."
        .Font.Bold = True : .Font.Name = "Arial" : .Font.Size = 10
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        .Interior.Color = RGB(30, 60, 120)
        .Font.Color = RGB(255, 255, 255)
    End With
    wsA.Columns("A").ColumnWidth = 5

    ' Sequenznummern in alle Datenzeilen schreiben (auch versteckte)
    Dim lastRow As Long : lastRow = wsA.Cells(wsA.Rows.Count, cArt).End(xlUp).Row
    Dim nr As Long : nr = 1
    Dim i As Long
    For i = ART_DATA_START To lastRow
        If Trim(CStr(wsA.Cells(i, cArt).Value)) <> "" Then
            wsA.Cells(i, 1).Value = nr
            wsA.Cells(i, 1).HorizontalAlignment = xlCenter
            nr = nr + 1
        Else
            wsA.Cells(i, 1).ClearContents
        End If
    Next i

    ' EAN-Spalte (B) sichtbar machen
    If wsA.Columns("B").ColumnWidth < 10 Then wsA.Columns("B").ColumnWidth = 16

    Application.EnableEvents   = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  MIGRATION: Leere Randspalte A einfügen
'  Verschiebt alle Daten um 1 Spalte nach rechts
'  Spalte A bleibt danach dauerhaft leer
' ================================================================
Sub Artikel_RandEinfuegen()
    Dim wsA As Worksheet : Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then MsgBox "Artikel-Sheet nicht gefunden!" : Exit Sub

    ' Prüfen ob Spalte A im Datenbereich bereits leer ist
    Dim testVal As String : testVal = Trim(CStr(wsA.Cells(ART_HEADER_ROW, 1).Value))
    If testVal = "" Or testVal = "Suche:" Then
        MsgBox "Spalte A ist bereits leer (kein Header in A" & ART_HEADER_ROW & ")." & Chr(10) & _
               "Migration wurde abgebrochen.", vbInformation, "Hinweis"
        Exit Sub
    End If

    If MsgBox("Spalte A wird als leerer Rand eingefügt." & Chr(10) & _
              "Alle Daten rücken eine Spalte nach rechts." & Chr(10) & Chr(10) & _
              "Bitte vorher Strg+S speichern!" & Chr(10) & Chr(10) & _
              "Jetzt fortfahren?", vbYesNo + vbQuestion, "Migration") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Erst alle versteckten Zeilen einblenden (alte Filter aufheben)
    Dim unR As Long
    For unR = ART_DATA_START To wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row + 100
        If wsA.Rows(unR).Hidden Then wsA.Rows(unR).Hidden = False
    Next unR

    ' Neue leere Spalte A einfügen → alles rückt nach rechts
    wsA.Columns(1).Insert Shift:=xlToRight

    ' Spalte A komplett leeren + schmal machen
    wsA.Columns(1).ClearContents
    wsA.Columns(1).ClearFormats
    wsA.Columns(1).ColumnWidth = 2
    wsA.Columns(1).Interior.ColorIndex = xlNone

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Setup neu aufrufen → repariert Titelzeile, Suchfeld, Buttons
    Setup_Artikelblatt

    MsgBox "Fertig! Spalte A ist jetzt leer." & Chr(10) & Chr(10) & _
           "Bitte prüfen:" & Chr(10) & _
           "  - Spaltenköpfe in Zeile 4 noch vorhanden?" & Chr(10) & _
           "  - Daten ab Zeile 5 sichtbar?" & Chr(10) & Chr(10) & _
           "Dann: Artikelblatt_FelderErgaenzen ausführen" & Chr(10) & _
           "Dann: Strg+S speichern.", vbInformation, "Migration abgeschlossen"
End Sub


' ================================================================
'  BEWEGUNGSHISTORIE - POPUP ANZEIGEN
' ================================================================
Sub Bewegungen_Zeigen()
    Dim wsD As Worksheet : Set wsD = GetSheet("ArtikelDetail")
    If wsD Is Nothing Then Exit Sub
    Dim ean As String : ean = CStr(wsD.Cells(4, 6).Value)
    If ean = "" Then MsgBox "Kein Artikel geoeffnet.", vbInformation : Exit Sub
    Dim wsB As Worksheet : Set wsB = GetSheet("Abg")
    If wsB Is Nothing Then
        MsgBox "Bewegungsblatt (Abg) nicht gefunden.", vbExclamation : Exit Sub
    End If
    Dim wsP As Worksheet : Set wsP = GetSheet("BewPopup")
    If wsP Is Nothing Then
        Setup_Bewegungen
        Set wsP = GetSheet("BewPopup")
        If wsP Is Nothing Then Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim lastOld As Long : lastOld = wsP.Cells(wsP.Rows.Count, 2).End(xlUp).Row
    If lastOld >= 4 Then
        wsP.Rows("4:" & lastOld).ClearContents
        wsP.Rows("4:" & lastOld).Interior.ColorIndex = xlNone
    End If
    wsP.Cells(2, 2).Value = CStr(wsD.Cells(3, 3).Value) & "  [EAN: " & ean & "]"
    Dim sRow As Long : sRow = 4
    Dim i As Long
    Dim lastBRow As Long : lastBRow = wsB.Cells(wsB.Rows.Count, 2).End(xlUp).Row
    For i = 2 To lastBRow
        If CStr(wsB.Cells(i, 2).Value) = ean Then
            wsP.Cells(sRow, 2).Value = wsB.Cells(i, 1).Value
            wsP.Cells(sRow, 2).NumberFormat = "DD.MM.YYYY HH:MM"
            wsP.Cells(sRow, 3).Value = wsB.Cells(i, 5).Value
            wsP.Cells(sRow, 4).Value = wsB.Cells(i, 6).Value
            wsP.Cells(sRow, 5).Value = wsB.Cells(i, 7).Value
            wsP.Cells(sRow, 6).Value = wsB.Cells(i, 8).Value
            If InStr(1, CStr(wsB.Cells(i, 6).Value), "Zugang", vbTextCompare) > 0 Then
                wsP.Cells(sRow, 4).Font.Color = RGB(0, 100, 0)
            Else
                wsP.Cells(sRow, 4).Font.Color = RGB(150, 0, 0)
            End If
            sRow = sRow + 1
        End If
    Next i
    If sRow = 4 Then wsP.Cells(4, 2).Value = "(keine Bewegungen gefunden)"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    wsP.Visible = xlSheetVisible
    wsP.Activate
    wsP.Cells(4, 1).Select
End Sub

' ================================================================
'  BEWEGUNGSHISTORIE - HANDLER
' ================================================================
Sub Bewegungen_Handler(ByVal Target As Range)
    If Target.Row = 2 Then
        Dim btn As String : btn = UCase(Trim(CStr(Target.Cells(1, 1).Value)))
        If btn = "SCHLIESSEN" Then Bewegungen_Schliessen
    End If
End Sub

' ================================================================
'  BEWEGUNGSHISTORIE - SCHLIESSEN
' ================================================================
Sub Bewegungen_Schliessen()
    Dim wsP As Worksheet : Set wsP = GetSheet("BewPopup")
    If Not wsP Is Nothing Then wsP.Visible = xlSheetHidden
    Dim wsD As Worksheet : Set wsD = GetSheet("ArtikelDetail")
    If Not wsD Is Nothing Then wsD.Activate
End Sub

' ================================================================
'  SETUP: BewPopup-Blatt erstellen
' ================================================================
Sub Setup_Bewegungen()
    Dim wsP As Worksheet : Set wsP = GetSheet("BewPopup")
    If wsP Is Nothing Then
        Set wsP = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsP.Name = "BewPopup"
    End If
    Application.DisplayAlerts = False
    wsP.Visible = xlSheetVisible
    wsP.Cells.UnMerge : wsP.Cells.ClearContents : wsP.Cells.ClearFormats
    Application.DisplayAlerts = True
    wsP.Cells.Font.Name = "Arial" : wsP.Cells.Font.Size = 11
    wsP.Columns("A").ColumnWidth = 2
    wsP.Columns("B").ColumnWidth = 18
    wsP.Columns("C").ColumnWidth = 10
    wsP.Columns("D").ColumnWidth = 18
    wsP.Columns("E").ColumnWidth = 16
    wsP.Columns("F").ColumnWidth = 14
    wsP.Rows(1).RowHeight = 30
    wsP.Rows(2).RowHeight = 22
    wsP.Rows(3).RowHeight = 20
    wsP.Range("A1:F1").Merge
    With wsP.Cells(1, 1)
        .Value = "Bewegungshistorie"
        .Interior.Color = RGB(0, 80, 160) : .Font.Color = RGB(255, 255, 255)
        .Font.Size = 14 : .Font.Bold = True
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    wsP.Range("B2:E2").Merge
    wsP.Cells(2, 2).Font.Bold = True : wsP.Cells(2, 2).Font.Size = 11
    With wsP.Cells(2, 6)
        .Value = "SCHLIESSEN"
        .Font.Bold = True : .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous : .Borders.Weight = xlMedium
    End With
    Dim hdrs As Variant
    hdrs = Array("", "Datum/Zeit", "Menge", "Typ", "Lagerort", "Benutzer")
    Dim col As Long
    For col = 1 To 6
        With wsP.Cells(3, col)
            .Value = hdrs(col - 1)
            .Font.Bold = True : .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(64, 64, 64)
            .HorizontalAlignment = xlCenter
        End With
    Next col
    wsP.Activate
    wsP.Cells(4, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    wsP.Visible = xlSheetHidden
    MsgBox "BewPopup-Sheet erstellt!" & Chr(10) & Chr(10) & _
           "Sheet-Modul-Code einfuegen (Alt+F11 -> BewPopup):" & Chr(10) & Chr(10) & _
           "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10) & _
           "    Bewegungen_Handler Target" & Chr(10) & "End Sub" & Chr(10) & Chr(10) & _
           "Strg+S speichern.", vbInformation, "Setup_Bewegungen"
End Sub
