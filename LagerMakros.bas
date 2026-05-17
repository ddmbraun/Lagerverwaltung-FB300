Attribute VB_Name = "LagerMakros"

' Einstellungen
Const ZEBRA_DRUCKER   As String = "ZDesigner GK420d"
Public Const BENUTZER        As String = "Frank"
Const INV_DATEN_START As Long = 6       ' Inventur: Datenzeilen ab Zeile 6

' Letzte angeklickte Artikelzeile
Public g_LetzteZeile As Long
' Gewaehlter Artikel in InvSuche
Public g_InvSucheArtikelZeile As Long

' ================================================================
'  SHEET SUCHEN (robust, ohne Umlaut-Probleme)
' ================================================================
Function GetSheet(suchbegriff As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, suchbegriff, vbTextCompare) > 0 Then
            Set GetSheet = ws
            Exit Function
        End If
    Next ws
    Set GetSheet = Nothing
End Function


' ================================================================
'  TOOLBAR-HANDLER
' ================================================================
Sub Toolbar_Handler(ByVal Target As Range)
    ' Neues Layout: Zeile 1=Titel, 2=Suche, 3=Toolbar, 4=Headers, 5+=Daten
    Dim ws As Worksheet: Set ws = Target.Worksheet
    Dim r As Long: r = Target.Row
    Dim col As Long: col = Target.Column

    ' Artikel-Zeilen ab Zeile 5: Auswahl markieren
    If r >= 5 Then
        If g_LetzteZeile >= 5 Then
            ws.Rows(g_LetzteZeile).Interior.ColorIndex = xlNone
        End If
        g_LetzteZeile = r
        ws.Rows(r).Interior.Color = RGB(255, 255, 153)
        Exit Sub
    End If

    ' Zeile 2: Suche-Buttons (G2=SUCHEN, H2=LEEREN)
    If r = 2 Then
        If col < 7 Or col > 8 Then Exit Sub
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        If col = 7 Then Artikel_Suchen
        If col = 8 Then Artikel_Suche_Leeren
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Zeile 3: Toolbar-Buttons
    If r = 3 Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Select Case col
            Case 4:  ZuAbgang_Buchen
            Case 6:  Etikett_Drucken
            Case 8:  EK_Toggle
            Case 10: Filter_Loeschen
            Case 12: NeueModule.Schnellansicht_Oeffnen
            Case 14: NeueModule.GitHub_Export
        End Select
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

' ================================================================
'  ARTIKEL ZEILE MARKIEREN (aufgerufen aus Worksheet_SelectionChange)
' ================================================================
Sub Artikel_Zeile_Markieren(ByVal Target As Range)
    Dim ws As Worksheet: Set ws = Target.Worksheet
    If g_LetzteZeile >= 5 Then
        ws.Rows(g_LetzteZeile).Interior.ColorIndex = xlNone
    End If
    g_LetzteZeile = Target.Row
    ws.Rows(Target.Row).Interior.Color = RGB(255, 255, 153)
End Sub

' ================================================================
'  ZU-/ABGANG BUCHEN
' ================================================================
Sub ZuAbgang_Buchen()
    If g_LetzteZeile < 5 Then
        MsgBox "Bitte zuerst eine Artikelzeile anklicken.", vbInformation
        Exit Sub
    End If

    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsB As Worksheet: Set wsB = GetSheet("Best")
    Dim wsZ As Worksheet: Set wsZ = GetSheet("Abg")

    If wsA Is Nothing Or wsB Is Nothing Or wsZ Is Nothing Then
        MsgBox "Sheet nicht gefunden. Gefundene Sheets:" & Chr(10) & SheetListe(), vbCritical
        Exit Sub
    End If

    Dim zeile As Long: zeile = g_LetzteZeile

    Dim colEAN   As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt   As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colAnz   As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim colNr    As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colLager As Long: colLager = Spalte_Finden(wsA, "LAGERORT")
    Dim colVK    As Long: colVK = Spalte_Finden(wsA, "VK-PREIS")

    Dim ean      As String: ean = wsA.Cells(zeile, colEAN).Value
    Dim artikel  As String: artikel = wsA.Cells(zeile, colArt).Value
    Dim artNr    As String: artNr = wsA.Cells(zeile, colNr).Value
    Dim lagerort As String: lagerort = wsA.Cells(zeile, colLager).Value
    Dim aktuell  As Double: aktuell = Val(wsA.Cells(zeile, colAnz).Value)

    Dim eingabe As String
    eingabe = InputBox( _
        "Artikel:  " & artikel & Chr(10) & _
        "ArtNr:    " & artNr & Chr(10) & _
        "Aktueller Bestand: " & Format(aktuell, "0") & " Stk" & Chr(10) & Chr(10) & _
        "Menge eingeben:" & Chr(10) & _
        "  Zugang  ->  positive Zahl  (z.B.  5)" & Chr(10) & _
        "  Abgang  ->  negative Zahl  (z.B.  -3)", _
        "Zu-/Abgang buchen")

    If eingabe = "" Then Exit Sub
    Dim menge As Double: menge = Val(eingabe)
    If menge = 0 Then MsgBox "Ungueltige Eingabe.": Exit Sub

    Dim typ As String: typ = IIf(menge > 0, "Zugang", "Abgang")
    Dim neuerBestand As Double: neuerBestand = aktuell + menge

    ' Artikel-Sheet aktualisieren
    wsA.Cells(zeile, colAnz).Value = neuerBestand

    ' Bewegung eintragen
    Dim nRow As Long
    nRow = wsZ.Cells(wsZ.Rows.Count, 1).End(xlUp).Row + 1
    wsZ.Cells(nRow, 1).Value = Now()
    wsZ.Cells(nRow, 1).NumberFormat = "DD.MM.YYYY HH:MM"
    wsZ.Cells(nRow, 2).Value = ean
    wsZ.Cells(nRow, 3).Value = artNr
    wsZ.Cells(nRow, 4).Value = artikel
    wsZ.Cells(nRow, 5).Value = Abs(menge)
    wsZ.Cells(nRow, 6).Value = typ
    wsZ.Cells(nRow, 7).Value = lagerort
    wsZ.Cells(nRow, 8).Value = BENUTZER

    ' Bestaende-Sheet aktualisieren
    Dim i As Long
    For i = 2 To wsB.Cells(wsB.Rows.Count, 2).End(xlUp).Row
        If wsB.Cells(i, 2).Value = artNr Then
            wsB.Cells(i, 4).Value = neuerBestand
            wsB.Cells(i, 6).Value = Round(neuerBestand * wsA.Cells(zeile, colVK).Value, 2)
            wsB.Cells(i, 10).Value = IIf(neuerBestand = 0, "! Nachbestellung", "OK")
            Exit For
        End If
    Next i

    MsgBox "OK - " & typ & ": " & Format(Abs(menge), "0") & " Stk" & Chr(10) & _
           "Neuer Bestand: " & Format(neuerBestand, "0") & " Stk", vbInformation
End Sub

' ================================================================
'  ETIKETT DRUCKEN
' ================================================================
Sub Etikett_Drucken()
    If g_LetzteZeile < 5 Then
        MsgBox "Bitte zuerst eine Artikelzeile anklicken.", vbInformation
        Exit Sub
    End If

    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim zeile As Long: zeile = g_LetzteZeile

    Dim colEAN   As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt   As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colVK    As Long: colVK = Spalte_Finden(wsA, "VK-PREIS")
    Dim colTextB As Long: colTextB = Spalte_Finden(wsA, "TextB")

    Dim ean     As String: ean = Trim(wsA.Cells(zeile, colEAN).Value)
    Dim artikel As String: artikel = Trim(wsA.Cells(zeile, colArt).Value)
    Dim vkPreis As String: vkPreis = Format(wsA.Cells(zeile, colVK).Value, "0.00") & " EUR"
    Dim textB   As String
    If colTextB > 0 Then textB = Trim(wsA.Cells(zeile, colTextB).Value)

    If ean = "" Then
        MsgBox "Kein EAN vorhanden.", vbExclamation
        Exit Sub
    End If

    ' Layout aus Taylor 70x38mm.lbl (203dpi = 8 dots/mm)
    ' Artikelbez: Y=1.13mm, TextB: Y=6.16mm, VKPreis: Y=14.7mm, Barcode: Y=19.65mm
    Dim zpl As String
    zpl = "^XA" & Chr(10)
    zpl = zpl & "^MMT^PW560^LL304^LS0" & Chr(10)
    zpl = zpl & "^FT12,49^A0N,38,38^FH\^FD" & Left(artikel, 35) & "^FS" & Chr(10)
    If Len(artikel) > 35 Then
        zpl = zpl & "^FT12,88^A0N,35,35^FH\^FD" & Mid(artikel, 36, 35) & "^FS" & Chr(10)
    ElseIf textB <> "" Then
        zpl = zpl & "^FT12,95^A0N,38,38^FH\^FD" & Left(textB, 35) & "^FS" & Chr(10)
    End If
    zpl = zpl & "^FT12,155^A0N,34,34^FH\^FD" & vkPreis & "^FS" & Chr(10)
    zpl = zpl & "^FT35,295^BCN,80,Y,N,N^FD" & ean & "^FS" & Chr(10)
    zpl = zpl & "^PQ1^XZ"

    Dim tmpDatei As String: tmpDatei = Environ("TEMP") & "\zebra_label.zpl"
    Dim ff As Integer: ff = FreeFile
    Open tmpDatei For Output As #ff
    Print #ff, zpl
    Close #ff

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    Dim ret As Long
    ret = oShell.Run("cmd /c copy /b """ & tmpDatei & """ """ & ZEBRA_DRUCKER & """", 0, True)
    If ret <> 0 Then
        MsgBox "Druckfehler (Code " & ret & ")" & Chr(10) & _
               "Druckername pruefen: '" & ZEBRA_DRUCKER & "'" & Chr(10) & Chr(10) & _
               "Windows: Einstellungen -> Drucker & Scanner" & Chr(10) & _
               "Dort den genauen Namen des Zebra-Druckers pruefen.", vbExclamation, "Druckfehler"
    Else
        MsgBox "Etikett gesendet an: " & ZEBRA_DRUCKER & Chr(10) & _
               "Artikel: " & artikel & Chr(10) & _
               "EAN: " & ean, vbInformation, "Etikett gedruckt"
    End If
End Sub

' ================================================================
'  EK TOGGLE
' ================================================================
Sub EK_Toggle()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim ekSpalte As Long: ekSpalte = Spalte_Finden(wsA, "EK-PREIS")
    If ekSpalte = 0 Then Exit Sub
    wsA.Columns(ekSpalte).Hidden = Not wsA.Columns(ekSpalte).Hidden
    wsA.Cells(3, 8).Value = IIf(wsA.Columns(ekSpalte).Hidden, "EK einbl.", "EK ausbl.")
End Sub

' ================================================================
'  FILTER LOESCHEN
' ================================================================
' ================================================================
'  ARTIKEL SUCHEN / SUCHE LEEREN (Suchfeld D2)
' ================================================================
Sub Artikel_Suchen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Dim such As String: such = Trim(wsA.Cells(2, 2).Value)
    If such = "" Then Artikel_Suche_Leeren: Exit Sub

    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colNr  As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    If colArt = 0 Then Exit Sub

    Dim lastRow As Long: lastRow = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
    Dim woerter() As String: woerter = Split(LCase(such), " ")
    Application.ScreenUpdating = False

    Dim treffer As Long: treffer = 0
    Dim i As Long, w As Integer, passt As Boolean, suchIn As String
    For i = 5 To lastRow
        suchIn = LCase(wsA.Cells(i, colArt).Value)
        If colEAN > 0 Then suchIn = suchIn & " " & LCase(wsA.Cells(i, colEAN).Value)
        If colNr > 0 Then suchIn = suchIn & " " & LCase(wsA.Cells(i, colNr).Value)
        passt = True
        For w = 0 To UBound(woerter)
            If Trim(woerter(w)) <> "" Then
                If InStr(suchIn, Trim(woerter(w))) = 0 Then passt = False: Exit For
            End If
        Next w
        wsA.Rows(i).Hidden = Not passt
        If passt Then treffer = treffer + 1
    Next i
    On Error Resume Next
    wsA.Shapes("trefferAnzeige").TextFrame.Characters.Text = treffer & " Treffer"
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Sub Artikel_Suche_Leeren()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Application.EnableEvents = False
    wsA.Cells(2, 2).Value = ""
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    If wsA.AutoFilterMode Then wsA.AutoFilter.ShowAllData
    ' UsedRange statt End(xlUp) – funktioniert auch wenn ALLE Zeilen versteckt sind
    Dim lastRow As Long
    With wsA.UsedRange
        lastRow = .Row + .Rows.Count - 1
    End With
    If lastRow < 5 Then lastRow = 5000
    wsA.Rows("5:" & lastRow).Hidden = False
    On Error Resume Next
    wsA.Shapes("trefferAnzeige").TextFrame.Characters.Text = "Treffer"
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  AKTUALISIEREN (Suchfeld leeren + alle Zeilen einblenden)
' ================================================================
Sub Artikel_Aktualisieren()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Application.EnableEvents = False
    wsA.Cells(2, 2).Value = ""
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    If wsA.AutoFilterMode Then wsA.AutoFilter.ShowAllData
    Dim lastRow As Long
    With wsA.UsedRange
        lastRow = .Row + .Rows.Count - 1
    End With
    If lastRow < 5 Then lastRow = 5000
    wsA.Rows("5:" & lastRow).Hidden = False
    On Error Resume Next
    wsA.Shapes("trefferAnzeige").TextFrame.Characters.Text = "Treffer"
    On Error GoTo 0
    Application.ScreenUpdating = True
    wsA.Cells(2, 2).Select
End Sub

Sub Filter_Loeschen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    ' Markierung entfernen
    If g_LetzteZeile >= 5 Then
        wsA.Rows(g_LetzteZeile).Interior.ColorIndex = xlNone
    End If
    g_LetzteZeile = 0
    Application.EnableEvents = False
    wsA.Cells(2, 2).Value = ""
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    If wsA.AutoFilterMode Then wsA.AutoFilter.ShowAllData
    ' Alle manuell versteckten Zeilen wieder einblenden
    Dim lastRow As Long
    With wsA.UsedRange
        lastRow = .Row + .Rows.Count - 1
    End With
    If lastRow < 5 Then lastRow = 5000
    wsA.Rows("5:" & lastRow).Hidden = False
    On Error Resume Next
    wsA.Shapes("trefferAnzeige").TextFrame.Characters.Text = "Treffer"
    On Error GoTo 0
    Application.ScreenUpdating = True
    wsA.Cells(5, 1).Select
End Sub

' ================================================================
'  HILFSFUNKTIONEN
' ================================================================
Function Spalte_Finden(ws As Worksheet, headerName As String) As Long
    Dim hRow As Long
    hRow = IIf(InStr(ws.Name, "rtikel") > 0, 4, 1)  ' Artikel: Header jetzt in Zeile 4
    Dim lastCol As Long
    lastCol = ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column
    Dim i As Long
    For i = 1 To lastCol
        If InStr(1, ws.Cells(hRow, i).Value, headerName, vbTextCompare) > 0 Then
            Spalte_Finden = i
            Exit Function
        End If
    Next i
    Spalte_Finden = 0
End Function

' ================================================================
'  SCHNELLANSICHT - HANDLER (Buttons in Zeile 2)
' ================================================================
Sub Schnellansicht_Handler(ByVal Target As Range)
    If Target.Row <> 2 Then Exit Sub
    ' Buttons: D2=4(SUCHEN), E2=5(FILTER), F2=6(AKTUALISIEREN),
    '          G2=7(EK AUSBL.), J2=10(SCHLIESSEN)
    Dim col As Long: col = Target.Column
    If col < 4 Or col > 10 Then Exit Sub
    If col = 8 Or col = 9 Then Exit Sub  ' H2/I2 = Treffer-Anzeige, kein Button
    On Error GoTo Fehler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Select Case col
        Case 4:  Schnellansicht_Suchen
        Case 5:  Schnellansicht_FilterLoeschen
        Case 6:  Schnellansicht_Aktualisieren
        Case 7:  NeueModule.Schnellansicht_EK_Toggle
        Case 10: NeueModule.Schnellansicht_Schliessen
    End Select
Fehler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  SCHNELLANSICHT SUCHEN + POPUP (mit Mehrwort-Suche)
' ================================================================
Sub Schnellansicht_Suchen()
    ' Neue Spalten: D=Artikel(4), E=EAN(5), C=Art.-Nr.(3)
    ' Treffer-Anzeige: H2(8)
    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub

    Dim such As String: such = Trim(wsS.Cells(2, 2).Value)

    ' Leer = alles anzeigen
    If such = "" Then
        Schnellansicht_FilterLoeschen
        Exit Sub
    End If

    ' Suchbegriffe aufteilen (Leerzeichen = AND-Verknüpfung)
    Dim woerter() As String: woerter = Split(LCase(such), " ")

    ' In Schnellansicht suchen: D=Artikel, C=Art.-Nr., E=EAN
    Dim lastSvRow As Long: lastSvRow = wsS.Cells(wsS.Rows.Count, 4).End(xlUp).Row
    Dim treffer As Long: treffer = 0
    Dim gefSvZeile As Long: gefSvZeile = 0

    ' EAN-Suche: wenn Suchbegriff nur aus Ziffern besteht (Val() hat Probleme mit langen EANs)
    Dim nurZahlen As Boolean: nurZahlen = False
    If Len(such) >= 2 Then
        Dim alleZiffern As Boolean: alleZiffern = True
        Dim ci As Integer
        For ci = 1 To Len(such)
            If InStr("0123456789", Mid(such, ci, 1)) = 0 Then alleZiffern = False: Exit For
        Next ci
        nurZahlen = alleZiffern
    End If

    Dim i As Long, w As Integer, passt As Boolean
    For i = 4 To lastSvRow
        If wsS.Rows(i).Hidden Then GoTo WeiterI  ' versteckte Zeilen ueberspringen beim Zaehlen nicht - wir zaehlen alle
        Dim suchIn As String
        If nurZahlen Then
            ' Nur Art.-Nr. + EAN
            suchIn = LCase(wsS.Cells(i, 3).Value & " " & wsS.Cells(i, 5).Value)
        Else
            ' Nur Artikeltext
            suchIn = LCase(wsS.Cells(i, 4).Value)
        End If
        passt = True
        For w = 0 To UBound(woerter)
            If Trim(woerter(w)) <> "" Then
                If InStr(suchIn, Trim(woerter(w))) = 0 Then
                    passt = False: Exit For
                End If
            End If
        Next w
        If passt Then
            treffer = treffer + 1
            If treffer = 1 Then gefSvZeile = i
        End If
WeiterI:
    Next i

    ' Treffer-Anzeige aktualisieren
    wsS.Cells(2, 8).Value = treffer & " Treffer"

    ' Genau 1 Treffer: Popup (alle Daten aus Schnellansicht)
    If treffer = 1 Then
        Dim anz As Double: anz = Val(wsS.Cells(gefSvZeile, 8).Value)  ' H=Bestand
        Dim einheit As String: einheit = wsS.Cells(gefSvZeile, 9).Value  ' I=Einheit
        Dim bestand As String
        If anz = 0 Then
            bestand = "0  !! NACHBESTELLUNG !!"
        ElseIf anz <= 5 Then
            bestand = Format(anz, "0") & " " & einheit & "  (Bestand niedrig!)"
        Else
            bestand = Format(anz, "0") & " " & einheit
        End If
        MsgBox wsS.Cells(gefSvZeile, 4).Value & Chr(10) & _
               String(40, "-") & Chr(10) & _
               "EAN:         " & wsS.Cells(gefSvZeile, 5).Value & Chr(10) & _
               "Art.-Nr.:    " & wsS.Cells(gefSvZeile, 3).Value & Chr(10) & _
               "VK-Preis:    " & Format(wsS.Cells(gefSvZeile, 6).Value, "0.00") & " EUR" & Chr(10) & _
               "Bestand:     " & bestand & Chr(10) & _
               "Lagerort:    " & wsS.Cells(gefSvZeile, 10).Value & Chr(10) & _
               "Warengruppe: " & wsS.Cells(gefSvZeile, 11).Value, _
               vbInformation, "Artikel gefunden"
        Exit Sub
    End If

    ' Kein Treffer
    If treffer = 0 Then
        MsgBox "Kein Artikel gefunden für: """ & such & """", vbExclamation, "Suche"
        Exit Sub
    End If

    ' Mehrere Treffer: Zeilen filtern
    Application.ScreenUpdating = False
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False
    Dim lastSvRow2 As Long: lastSvRow2 = wsS.Cells(wsS.Rows.Count, 4).End(xlUp).Row
    Dim zSuch As String
    Dim zPasst As Boolean
    Dim j As Long
    For j = 4 To lastSvRow2
        If nurZahlen Then
            zSuch = LCase(wsS.Cells(j, 4).Value & " " & wsS.Cells(j, 5).Value & " " & wsS.Cells(j, 3).Value)
        Else
            zSuch = LCase(wsS.Cells(j, 4).Value & " " & wsS.Cells(j, 3).Value & " " & wsS.Cells(j, 10).Value & " " & wsS.Cells(j, 11).Value)
        End If
        zPasst = True
        For w = 0 To UBound(woerter)
            If Trim(woerter(w)) <> "" Then
                If InStr(zSuch, Trim(woerter(w))) = 0 Then
                    zPasst = False: Exit For
                End If
            End If
        Next w
        wsS.Rows(j).Hidden = Not zPasst
    Next j
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  SCHNELLANSICHT FILTER LOESCHEN
' ================================================================
Sub Schnellansicht_FilterLoeschen()
    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False
    Dim lastRow As Long
    lastRow = wsS.Cells(wsS.Rows.Count, 4).End(xlUp).Row  ' Artikel-Spalte D
    If lastRow >= 4 Then wsS.Rows("4:" & lastRow).Hidden = False
    wsS.Cells(2, 2).Value = ""      ' Suchfeld leeren
    wsS.Cells(2, 8).Value = ""      ' Treffer-Anzeige leeren
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  SCHNELLANSICHT AKTUALISIEREN
' ================================================================
Sub Schnellansicht_Aktualisieren()
    ' Neue Spalten: A=Spacer, B=#, C=Art.-Nr., D=Artikel, E=EAN,
    '               F=VK-Preis, G=EK-Preis, H=Bestand, I=Einheit,
    '               J=Lagerort, K=Warengruppe, L=Attribut, M=Zeilenverweis(hidden)
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsA Is Nothing Or wsS Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Alle Zeilen einblenden, dann loeschen
    Dim lastSvRow As Long
    lastSvRow = wsS.Cells(wsS.Rows.Count, 4).End(xlUp).Row  ' Artikel-Spalte D
    If lastSvRow < 4 Then lastSvRow = wsS.UsedRange.Rows.Count
    If lastSvRow >= 4 Then
        wsS.Rows("4:" & lastSvRow).Hidden = False
        wsS.Range("A4:M" & lastSvRow).ClearContents
    End If

    ' Spalten im Artikel-Sheet ermitteln
    Dim colEAN  As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt  As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr   As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colVK   As Long: colVK = Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK   As Long: colEK = Spalte_Finden(wsA, "EK-PREIS")
    Dim colAnz  As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim colEinh As Long: colEinh = Spalte_Finden(wsA, "EINHEIT")
    Dim colLag  As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colWG   As Long: colWG = Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colAttr As Long: colAttr = Spalte_Finden(wsA, "ATTRIBUT")

    If colArt = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Spalte 'ARTIKEL' nicht gefunden! Bitte Spaltenüberschrift prüfen.", vbCritical
        Exit Sub
    End If

    Dim lastRow As Long: lastRow = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
    Dim sRow As Long: sRow = 4
    Dim i As Long
    For i = 5 To lastRow
        If wsA.Cells(i, colArt).Value <> "" Then
            wsS.Cells(sRow, 1).Value = ""                                                              ' A: Spacer
            wsS.Cells(sRow, 2).Value = sRow - 3                                                        ' B: #
            wsS.Cells(sRow, 3).NumberFormat = "@"                                                      ' C: Art.-Nr. als Text
            If colNr > 0 Then wsS.Cells(sRow, 3).Value = CStr(wsA.Cells(i, colNr).Value)              ' C: Art.-Nr.
            wsS.Cells(sRow, 4).Value = wsA.Cells(i, colArt).Value                                     ' D: Artikel
            wsS.Cells(sRow, 5).NumberFormat = "@"                                                      ' E: EAN als Text
            If colEAN > 0 Then wsS.Cells(sRow, 5).Value = CStr(wsA.Cells(i, colEAN).Value)            ' E: EAN
            If colVK > 0 Then wsS.Cells(sRow, 6).Value = wsA.Cells(i, colVK).Value                    ' F: VK-Preis
            If colEK > 0 Then wsS.Cells(sRow, 7).Value = wsA.Cells(i, colEK).Value                    ' G: EK-Preis
            If colAnz > 0 Then wsS.Cells(sRow, 8).Value = wsA.Cells(i, colAnz).Value                  ' H: Bestand
            If colEinh > 0 Then wsS.Cells(sRow, 9).Value = wsA.Cells(i, colEinh).Value                ' I: Einheit
            If colLag > 0 Then wsS.Cells(sRow, 10).Value = wsA.Cells(i, colLag).Value                 ' J: Lagerort
            If colWG > 0 Then wsS.Cells(sRow, 11).Value = wsA.Cells(i, colWG).Value                   ' K: Warengruppe
            If colAttr > 0 Then wsS.Cells(sRow, 12).Value = wsA.Cells(i, colAttr).Value               ' L: Attribut
            wsS.Cells(sRow, 13).Value = i                                                              ' M: Zeilenverweis
            sRow = sRow + 1
        End If
    Next i

    ' Treffer-Anzeige in H2
    Dim anzahl As Long: anzahl = sRow - 4
    wsS.Cells(2, 8).Value = anzahl & " Treffer"

    ' EK-Preis (Spalte G) standardmaessig ausblenden
    If anzahl > 0 Then
        wsS.Range(wsS.Cells(4, 7), wsS.Cells(sRow - 1, 7)).NumberFormat = ";;;"
    End If
    wsS.Cells(3, 7).Value = ""        ' Kopfzeile leeren
    wsS.Cells(2, 7).Value = "EK EINBL."  ' Button-Text anpassen

    Application.ScreenUpdating = True
    Application.StatusBar = "Schnellansicht: " & anzahl & " Artikel aktualisiert."
End Sub

' ================================================================
'  NEUER ARTIKEL
' ================================================================
Sub NeuerArtikel()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsB As Worksheet: Set wsB = GetSheet("Best")
    If wsA Is Nothing Then Exit Sub

    ' --- Eingabe ---
    Dim artName As String
    artName = InputBox("Artikelname (Pflichtfeld):", "Neuer Artikel 1/5")
    If Trim(artName) = "" Then Exit Sub

    Dim ean As String
    ean = InputBox("EAN13 (oder leer lassen):", "Neuer Artikel 2/5")

    Dim artNr As String
    artNr = InputBox("Artikelnummer:", "Neuer Artikel 2/5")

    Dim vkStr As String
    vkStr = InputBox("VK-Preis (z.B. 9.99):", "Neuer Artikel 3/5", "0.00")
    Dim vkPreis As Double: vkPreis = Val(Replace(vkStr, ",", "."))

    Dim ekStr As String
    ekStr = InputBox("EK-Preis (z.B. 5.00):", "Neuer Artikel 3/5", "0.00")
    Dim ekPreis As Double: ekPreis = Val(Replace(ekStr, ",", "."))

    Dim mwstStr As String
    mwstStr = InputBox("MwSt % (Standard: 19):", "Neuer Artikel 4/5", "19")
    Dim mwst As Double: mwst = Val(mwstStr)
    If mwst = 0 Then mwst = 19

    Dim anzStr As String
    anzStr = InputBox("Anfangsbestand:", "Neuer Artikel 4/5", "0")
    Dim anzahl As Double: anzahl = Val(anzStr)

    Dim einheit As String
    einheit = InputBox("Einheit (Stk / Pkg / m ...):", "Neuer Artikel 4/5", "Stk")

    Dim warengruppe As String
    warengruppe = InputBox("Warengruppe:", "Neuer Artikel 5/5")

    Dim lagerort As String
    lagerort = InputBox("Lagerort:", "Neuer Artikel 5/5")

    ' --- Spalten ermitteln ---
    Dim colEAN2   As Long: colEAN2 = Spalte_Finden(wsA, "EAN13")
    Dim colArt2   As Long: colArt2 = Spalte_Finden(wsA, "ARTIKEL")
    Dim colVK2    As Long: colVK2 = Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK2    As Long: colEK2 = Spalte_Finden(wsA, "EK-PREIS")
    Dim colMwst2  As Long: colMwst2 = Spalte_Finden(wsA, "MWST")
    Dim colAnz2   As Long: colAnz2 = Spalte_Finden(wsA, "ANZAHL")
    Dim colEinh2  As Long: colEinh2 = Spalte_Finden(wsA, "EINHEIT")
    Dim colNr2    As Long: colNr2 = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colWG2    As Long: colWG2 = Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colLager2 As Long: colLager2 = Spalte_Finden(wsA, "LAGERORT")

    ' --- Neue Zeile einfuegen ---
    Dim nRow As Long
    nRow = wsA.Cells(wsA.Rows.Count, colArt2).End(xlUp).Row + 1

    If colEAN2 > 0 Then wsA.Cells(nRow, colEAN2).Value = ean
    If colEAN2 > 0 Then wsA.Cells(nRow, colEAN2).NumberFormat = "@"
    If colArt2 > 0 Then wsA.Cells(nRow, colArt2).Value = artName
    If colNr2 > 0 Then wsA.Cells(nRow, colNr2).Value = artNr
    If colVK2 > 0 Then wsA.Cells(nRow, colVK2).Value = vkPreis
    If colEK2 > 0 Then wsA.Cells(nRow, colEK2).Value = ekPreis
    If colMwst2 > 0 Then wsA.Cells(nRow, colMwst2).Value = mwst
    If colAnz2 > 0 Then wsA.Cells(nRow, colAnz2).Value = anzahl
    If colEinh2 > 0 Then wsA.Cells(nRow, colEinh2).Value = einheit
    If colWG2 > 0 Then wsA.Cells(nRow, colWG2).Value = warengruppe
    If colLager2 > 0 Then wsA.Cells(nRow, colLager2).Value = lagerort

    ' --- Bestaende aktualisieren ---
    If Not wsB Is Nothing Then
        Dim bRow As Long
        bRow = wsB.Cells(wsB.Rows.Count, 3).End(xlUp).Row + 1
        wsB.Cells(bRow, 1).Value = ean
        wsB.Cells(bRow, 2).Value = artNr
        wsB.Cells(bRow, 3).Value = artName
        wsB.Cells(bRow, 4).Value = anzahl
        wsB.Cells(bRow, 5).Value = einheit
        wsB.Cells(bRow, 6).Value = Round(anzahl * vkPreis, 2)
        wsB.Cells(bRow, 10).Value = IIf(anzahl = 0, "! Nachbestellung", "OK")
    End If

    ' Zur neuen Zeile springen
    wsA.Cells(nRow, colArt2).Select
    MsgBox "Neuer Artikel angelegt: " & artName, vbInformation
End Sub

Sub Diagnose()
    Dim msg As String
    msg = "=== DIAGNOSE ===" & Chr(10) & Chr(10)

    ' EnableEvents prüfen
    msg = msg & "EnableEvents: " & Application.EnableEvents & Chr(10)

    ' Schnellansicht-Sheet prüfen
    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then
        msg = msg & "FEHLER: Schnellansicht-Sheet nicht gefunden!" & Chr(10)
    Else
        msg = msg & "Schnellansicht-Sheet: OK (" & wsS.Name & ")" & Chr(10)
        ' Code in Sheet prüfen
        Dim cm As Object
        Set cm = ThisWorkbook.VBProject.VBComponents(wsS.CodeName).CodeModule
        msg = msg & "Code-Zeilen in Tabelle17: " & cm.CountOfLines & Chr(10)
        If cm.CountOfLines > 0 Then
            msg = msg & "Erste Zeile: " & cm.Lines(1, 1) & Chr(10)
        End If
        ' Suchfeld-Wert
        msg = msg & "Suchfeld B2: """ & wsS.Cells(2, 2).Value & """" & Chr(10)
    End If

    ' Artikel-Sheet prüfen
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then
        msg = msg & "FEHLER: Artikel-Sheet nicht gefunden!" & Chr(10)
    Else
        msg = msg & "Artikel-Sheet: OK (" & wsA.Name & ")" & Chr(10)
    End If

    MsgBox msg, vbInformation, "Diagnose"

    ' EnableEvents reparieren falls nötig
    If Not Application.EnableEvents Then
        Application.EnableEvents = True
        MsgBox "EnableEvents war False - wurde repariert!", vbExclamation
    End If
End Sub

Function SheetListe() As String
    Dim ws As Worksheet
    Dim s As String
    For Each ws In ThisWorkbook.Sheets
        s = s & ws.Name & Chr(10)
    Next ws
    SheetListe = s
End Function

' ================================================================
'  SETUP - richtet Tabelle17 automatisch ein (einmalig ausführen)
' ================================================================
Sub Setup_Ausfuehren()
    ' --- Schnellansicht einrichten ---
    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then
        MsgBox "Schnellansicht-Sheet nicht gefunden!", vbCritical
        Exit Sub
    End If

    Dim vbS As Object: Set vbS = ThisWorkbook.VBProject.VBComponents(wsS.CodeName)
    Dim cmS As Object: Set cmS = vbS.CodeModule
    If cmS.CountOfLines > 0 Then cmS.DeleteLines 1, cmS.CountOfLines
    Dim c As String: c = ""
    c = c & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    c = c & "    If Target.Address = ""$B$2"" Then" & Chr(10)
    c = c & "        On Error GoTo Fehler" & Chr(10)
    c = c & "        Application.EnableEvents = False" & Chr(10)
    c = c & "        LagerMakros.Schnellansicht_Suchen" & Chr(10)
    c = c & "        Application.EnableEvents = True" & Chr(10)
    c = c & "        Exit Sub" & Chr(10)
    c = c & "Fehler: Application.EnableEvents = True" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "" & Chr(10)
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    If Target.Row = 2 And Target.Column = 5 Then" & Chr(10)
    c = c & "        Cancel = True" & Chr(10)
    c = c & "        LagerMakros.Schnellansicht_FilterLoeschen" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "    If Target.Row = 2 And Target.Column = 6 Then" & Chr(10)
    c = c & "        Cancel = True" & Chr(10)
    c = c & "        LagerMakros.Schnellansicht_Aktualisieren" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cmS.AddFromString c

    ' --- Artikel-Sheet einrichten ---
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If Not wsA Is Nothing Then
        Dim vbA As Object: Set vbA = ThisWorkbook.VBProject.VBComponents(wsA.CodeName)
        Dim cmA As Object: Set cmA = vbA.CodeModule
        If cmA.CountOfLines > 0 Then cmA.DeleteLines 1, cmA.CountOfLines
        Dim a As String: a = ""
        a = a & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10)
        a = a & "    LagerMakros.Toolbar_Handler Target" & Chr(10)
        a = a & "End Sub" & Chr(10)
        cmA.AddFromString a
    End If

    ' --- Inventur-Sheet erstellen ---
    Inventur_Setup True

    ' --- InvSuche-Sheet erstellen ---
    InvSuche_Setup True

    MsgBox "Setup fertig!" & Chr(10) & Chr(10) & _
           "Artikel-Sheet: Zeile anklicken = markieren, Buttons in Zeile 1 aktiv" & Chr(10) & _
           "Schnellansicht: Suchbegriff + ENTER, Doppelklick LOESCHEN / AKTUALISIEREN" & Chr(10) & _
           "Inventur-Sheet: Doppelklick BEFUELLEN / UEBERNEHMEN", _
           vbInformation, "Setup abgeschlossen"
End Sub

' ================================================================
'  INVENTUR - SHEET ERSTELLEN
' ================================================================
Sub Inventur_Setup(Optional silent As Boolean = False)
    Dim wsI As Worksheet: Set wsI = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "Inventur", vbTextCompare) > 0 Then
            Set wsI = ws: Exit For
        End If
    Next ws
    If wsI Is Nothing Then
        Set wsI = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsI.Name = "Inventur"
    End If

    Application.ScreenUpdating = False
    wsI.Cells.Clear
    wsI.Cells.Interior.ColorIndex = xlNone

    Dim blau As Long: blau = RGB(31, 56, 100)
    Dim hellblau As Long: hellblau = RGB(46, 80, 144)
    Dim gruen As Long: gruen = RGB(55, 110, 50)
    Dim orange As Long: orange = RGB(180, 90, 0)
    Dim hellgrau As Long: hellgrau = RGB(242, 242, 242)

    ' --- Zeile 1: Titel ---
    wsI.Range("A1:J1").Merge
    wsI.Cells(1, 1).Value = "INVENTURLISTE"
    wsI.Cells(1, 1).Interior.Color = blau
    wsI.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsI.Cells(1, 1).Font.Size = 14
    wsI.Cells(1, 1).Font.Bold = True
    wsI.Cells(1, 1).HorizontalAlignment = xlCenter
    wsI.Rows(1).RowHeight = 30

    ' --- Zeile 2: Datum + BEFUELLEN + UEBERNEHMEN ---
    wsI.Cells(2, 1).Value = "Datum:"
    wsI.Cells(2, 1).Font.Bold = True
    wsI.Cells(2, 2).Value = Date
    wsI.Cells(2, 2).NumberFormat = "DD.MM.YYYY"
    wsI.Range("B2:C2").Merge
    wsI.Cells(2, 4).Value = "Erstellt von:"
    wsI.Cells(2, 4).Font.Bold = True
    wsI.Cells(2, 5).Value = BENUTZER
    wsI.Range("E2:F2").Merge
    wsI.Range("G2:H2").Merge
    wsI.Cells(2, 7).Value = "BEFUELLEN"
    wsI.Cells(2, 7).Interior.Color = hellblau
    wsI.Cells(2, 7).Font.Color = RGB(255, 255, 255)
    wsI.Cells(2, 7).Font.Bold = True
    wsI.Cells(2, 7).HorizontalAlignment = xlCenter
    wsI.Range("I2:J2").Merge
    wsI.Cells(2, 9).Value = "UEBERNEHMEN"
    wsI.Cells(2, 9).Interior.Color = gruen
    wsI.Cells(2, 9).Font.Color = RGB(255, 255, 255)
    wsI.Cells(2, 9).Font.Bold = True
    wsI.Cells(2, 9).HorizontalAlignment = xlCenter
    wsI.Rows(2).RowHeight = 24

    ' --- Zeile 3: Suchfeld + Eingabe ---
    wsI.Cells(3, 1).Value = "Suche:"
    wsI.Cells(3, 1).Font.Bold = True
    wsI.Cells(3, 1).HorizontalAlignment = xlRight
    ' B3 = Suchfeld (EAN oder Name) - Text-Format damit EAN nicht als Zahl dargestellt wird
    wsI.Cells(3, 2).NumberFormat = "@"
    wsI.Cells(3, 2).Interior.Color = RGB(255, 255, 220)
    wsI.Cells(3, 2).Font.Bold = True
    wsI.Range("B3:D3").Merge
    ' E3:F3 = gefundener Artikelname
    wsI.Range("E3:G3").Merge
    wsI.Cells(3, 5).Interior.Color = hellgrau
    ' H3 = Soll-Anzeige
    wsI.Cells(3, 8).Interior.Color = hellgrau
    ' I3 = Gezaehlt-Eingabe
    wsI.Cells(3, 9).Interior.Color = RGB(220, 255, 220)
    wsI.Cells(3, 9).Font.Bold = True
    ' J3 = EINTRAGEN Button
    wsI.Cells(3, 10).Value = "EINTRAGEN"
    wsI.Cells(3, 10).Interior.Color = orange
    wsI.Cells(3, 10).Font.Color = RGB(255, 255, 255)
    wsI.Cells(3, 10).Font.Bold = True
    wsI.Cells(3, 10).HorizontalAlignment = xlCenter
    ' Labels in Zeile 4
    wsI.Cells(4, 5).Value = "Artikel"
    wsI.Cells(4, 5).Font.Size = 8
    wsI.Cells(4, 5).Font.Color = RGB(120, 120, 120)
    wsI.Cells(4, 8).Value = "Soll"
    wsI.Cells(4, 8).Font.Size = 8
    wsI.Cells(4, 8).Font.Color = RGB(120, 120, 120)
    wsI.Cells(4, 9).Value = "Menge"
    wsI.Cells(4, 9).Font.Size = 8
    wsI.Cells(4, 9).Font.Color = RGB(120, 120, 120)
    wsI.Rows(3).RowHeight = 26
    wsI.Rows(4).RowHeight = 14

    ' --- Zeile 5: Spaltenkoepfe Liste ---
    Dim hdr As Variant
    hdr = Array("Nr", "EAN", "Artikel", "Lagerort", "EK-Preis", "SOLL", "GEZAEHLT", "DIFFERENZ", "EK-Wert", "Bemerkung")
    Dim j As Integer
    For j = 0 To 9
        wsI.Cells(5, j + 1).Value = hdr(j)
        wsI.Cells(5, j + 1).Interior.Color = hellblau
        wsI.Cells(5, j + 1).Font.Color = RGB(255, 255, 255)
        wsI.Cells(5, j + 1).Font.Bold = True
        wsI.Cells(5, j + 1).HorizontalAlignment = xlCenter
    Next j
    wsI.Rows(5).RowHeight = 20

    ' --- Spaltenbreiten ---
    wsI.Columns(1).ColumnWidth = 5
    wsI.Columns(2).ColumnWidth = 16
    wsI.Columns(3).ColumnWidth = 32
    wsI.Columns(4).ColumnWidth = 14
    wsI.Columns(5).ColumnWidth = 10
    wsI.Columns(6).ColumnWidth = 8
    wsI.Columns(7).ColumnWidth = 10
    wsI.Columns(8).ColumnWidth = 12
    wsI.Columns(9).ColumnWidth = 12
    wsI.Columns(10).ColumnWidth = 18

    ' --- Sheet-Events installieren ---
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsI.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim c As String: c = ""
    c = c & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    c = c & "    If Target.Address = ""$B$3"" Then" & Chr(10)
    c = c & "        On Error GoTo Fehler" & Chr(10)
    c = c & "        Application.EnableEvents = False" & Chr(10)
    c = c & "        LagerMakros.Inventur_Suchen" & Chr(10)
    c = c & "        Application.EnableEvents = True" & Chr(10)
    c = c & "        Exit Sub" & Chr(10)
    c = c & "Fehler: Application.EnableEvents = True" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    Cancel = True" & Chr(10)
    c = c & "    If Target.Row = 2 And Target.Column >= 7 And Target.Column <= 8 Then LagerMakros.Inventur_Befuellen" & Chr(10)
    c = c & "    If Target.Row = 2 And Target.Column >= 9 Then LagerMakros.Inventur_BestaendeUebernehmen" & Chr(10)
    c = c & "    If Target.Row = 3 And Target.Column = 10 Then LagerMakros.Inventur_Eintragen" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c

    Application.ScreenUpdating = True
    If Not silent Then MsgBox "Inventur-Sheet erstellt!" & Chr(10) & Chr(10) & _
        "Bedienung:" & Chr(10) & _
        "- Suchfeld B3: EAN oder Artikelname eingeben + ENTER" & Chr(10) & _
        "- Menge in I3 eintragen" & Chr(10) & _
        "- Doppelklick EINTRAGEN = in Liste uebernehmen" & Chr(10) & _
        "- Doppelklick BEFUELLEN = alle Artikel laden" & Chr(10) & _
        "- Doppelklick UEBERNEHMEN = Bestaende aktualisieren", vbInformation
End Sub

' ================================================================
'  INVENTUR - ARTIKEL SUCHEN (via B3)
' ================================================================
Sub Inventur_Suchen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsI Is Nothing Then Exit Sub

    Dim such As String: such = Trim(wsI.Cells(3, 2).Value)
    ' Ergebnisfelder leeren
    wsI.Range("E3:G3").ClearContents
    wsI.Cells(3, 8).ClearContents

    If such = "" Then Exit Sub

    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    If colArt = 0 Then Exit Sub

    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
    Dim treffer As Long: treffer = 0
    Dim gefZeile As Long: gefZeile = 0
    Dim nurZahlen As Boolean: nurZahlen = (such = CStr(Val(such)) And Val(such) > 0)

    Dim i As Long
    For i = 5 To lastA
        Dim suchIn As String
        If nurZahlen Then
            suchIn = LCase(wsA.Cells(i, colArt).Value & " " & wsA.Cells(i, colEAN).Value)
        Else
            suchIn = LCase(wsA.Cells(i, colArt).Value)
        End If
        If InStr(suchIn, LCase(such)) > 0 Then
            treffer = treffer + 1
            If treffer = 1 Then gefZeile = i
        End If
    Next i

    If treffer = 1 Then
        wsI.Cells(3, 5).Value = wsA.Cells(gefZeile, colArt).Value
        If colAnz > 0 Then wsI.Cells(3, 8).Value = wsA.Cells(gefZeile, colAnz).Value
        wsI.Cells(3, 9).Select
    ElseIf treffer = 0 Then
        wsI.Cells(3, 5).Value = "-- Nicht gefunden --"
    Else
        wsI.Cells(3, 5).Value = treffer & " Treffer - genauer suchen"
    End If
End Sub

' ================================================================
'  INVENTUR - MENGE EINTRAGEN
' ================================================================
Sub Inventur_Eintragen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsI Is Nothing Then Exit Sub

    Dim such As String: such = Trim(wsI.Cells(3, 2).Value)
    Dim artName As String: artName = Trim(wsI.Cells(3, 5).Value)
    Dim mengeStr As String: mengeStr = Trim(wsI.Cells(3, 9).Value)

    If such = "" Or artName = "" Or artName = "-- Nicht gefunden --" Then
        MsgBox "Bitte zuerst einen Artikel suchen.", vbExclamation: Exit Sub
    End If
    If mengeStr = "" Then
        MsgBox "Bitte Menge eingeben.", vbExclamation: Exit Sub
    End If
    Dim menge As Double: menge = Val(mengeStr)

    ' Artikel in Liste (ab Zeile INV_DATEN_START) suchen
    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.Count, 3).End(xlUp).Row
    Dim gefunden As Boolean: gefunden = False
    Dim i As Long
    For i = INV_DATEN_START To lastI
        If LCase(Trim(wsI.Cells(i, 3).Value)) = LCase(artName) Then
            wsI.Cells(i, 7).Value = menge
            gefunden = True
            ' Zeile kurz hervorheben
            wsI.Cells(i, 7).Select
            Exit For
        End If
    Next i

    ' Nicht in Liste: neuen Eintrag hinzufuegen
    If Not gefunden Then
        Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
        Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
        Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
        Dim colEK  As Long: colEK = Spalte_Finden(wsA, "EK-PREIS")
        Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
        Dim lastA  As Long: lastA = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
        Dim nRow   As Long
        If lastI < INV_DATEN_START Then nRow = INV_DATEN_START Else nRow = lastI + 1
        Dim j As Long
        For j = 5 To lastA
            If LCase(Trim(wsA.Cells(j, colArt).Value)) = LCase(artName) Then
                wsI.Cells(nRow, 1).Value = nRow - INV_DATEN_START + 1
                If colEAN > 0 Then wsI.Cells(nRow, 2).Value = wsA.Cells(j, colEAN).Value
                wsI.Cells(nRow, 3).Value = wsA.Cells(j, colArt).Value
                If colLag > 0 Then wsI.Cells(nRow, 4).Value = wsA.Cells(j, colLag).Value
                If colEK > 0 Then wsI.Cells(nRow, 5).Value = wsA.Cells(j, colEK).Value
                If colAnz > 0 Then wsI.Cells(nRow, 6).Value = wsA.Cells(j, colAnz).Value
                wsI.Cells(nRow, 7).Value = menge
                wsI.Cells(nRow, 8).Formula = "=IF(G" & nRow & "="""","""",G" & nRow & "-F" & nRow & ")"
                wsI.Cells(nRow, 9).Formula = "=IF(G" & nRow & "<>"""",E" & nRow & "*G" & nRow & ",E" & nRow & "*F" & nRow & ")"
                wsI.Cells(nRow, 9).NumberFormat = "0.00"
                Exit For
            End If
        Next j
    End If

    ' Suchfelder leeren fuer naechsten Artikel
    Application.EnableEvents = False
    wsI.Cells(3, 2).ClearContents
    wsI.Range("E3:G3").ClearContents
    wsI.Cells(3, 8).ClearContents
    wsI.Cells(3, 9).ClearContents
    Application.EnableEvents = True
    wsI.Cells(3, 2).Select
End Sub

' ================================================================
'  INVENTUR - DATEN BEFUELLEN
' ================================================================
Sub Inventur_Befuellen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsI Is Nothing Then Exit Sub

    If MsgBox("Vorhandene Inventurdaten loeschen und neu laden?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.Count, 3).End(xlUp).Row
    If lastI >= INV_DATEN_START Then wsI.Range("A" & INV_DATEN_START & ":J" & lastI).Clear

    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colEK  As Long: colEK = Spalte_Finden(wsA, "EK-PREIS")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    If colArt = 0 Then Application.ScreenUpdating = True: Exit Sub

    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
    Dim sRow As Long: sRow = INV_DATEN_START
    Dim nr As Long: nr = 1
    Dim hellgrau As Long: hellgrau = RGB(242, 242, 242)

    Dim i As Long
    For i = 5 To lastA
        If wsA.Cells(i, colArt).Value <> "" Then
            wsI.Cells(sRow, 1).Value = nr
            If colEAN > 0 Then wsI.Cells(sRow, 2).Value = wsA.Cells(i, colEAN).Value
            wsI.Cells(sRow, 3).Value = wsA.Cells(i, colArt).Value
            If colLag > 0 Then wsI.Cells(sRow, 4).Value = wsA.Cells(i, colLag).Value
            If colEK > 0 Then
                wsI.Cells(sRow, 5).Value = wsA.Cells(i, colEK).Value
                wsI.Cells(sRow, 5).NumberFormat = "0.00"
            End If
            If colAnz > 0 Then wsI.Cells(sRow, 6).Value = wsA.Cells(i, colAnz).Value
            wsI.Cells(sRow, 8).Formula = "=IF(G" & sRow & "="""","""",G" & sRow & "-F" & sRow & ")"
            wsI.Cells(sRow, 9).Formula = "=IF(G" & sRow & "<>"""",E" & sRow & "*G" & sRow & ",E" & sRow & "*F" & sRow & ")"
            wsI.Cells(sRow, 9).NumberFormat = "0.00"
            If sRow Mod 2 = 0 Then wsI.Range("A" & sRow & ":J" & sRow).Interior.Color = hellgrau
            sRow = sRow + 1
            nr = nr + 1
        End If
    Next i

    ' Summenzeile
    Dim sumRow As Long: sumRow = sRow + 1
    wsI.Range("A" & sumRow & ":J" & sumRow).Interior.Color = RGB(31, 56, 100)
    wsI.Range("A" & sumRow & ":J" & sumRow).Font.Color = RGB(255, 255, 255)
    wsI.Range("A" & sumRow & ":J" & sumRow).Font.Bold = True
    wsI.Range("A" & sumRow & ":H" & sumRow).Merge
    wsI.Cells(sumRow, 1).Value = "SUMME EK-WERT (gezaehlt / gesamt):"
    wsI.Cells(sumRow, 1).HorizontalAlignment = xlRight
    wsI.Cells(sumRow, 9).Formula = "=SUM(I" & INV_DATEN_START & ":I" & (sRow - 1) & ")"
    wsI.Cells(sumRow, 9).NumberFormat = "#,##0.00 " & ChrW(8364)
    wsI.Rows(sumRow).RowHeight = 22

    Application.ScreenUpdating = True
    MsgBox "Inventur geladen: " & (nr - 1) & " Artikel.", vbInformation
End Sub

' ================================================================
'  INVENTUR - BESTAENDE UEBERNEHMEN
' ================================================================
Sub Inventur_BestaendeUebernehmen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsI Is Nothing Then Exit Sub

    If MsgBox("Bestaende im Artikel-Sheet mit den gezaehlten Mengen aktualisieren?", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    If colEAN = 0 Or colAnz = 0 Then Exit Sub

    Dim updated As Long: updated = 0
    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.Count, 3).End(xlUp).Row
    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.Count, colEAN).End(xlUp).Row

    Dim i As Long, j As Long
    For i = INV_DATEN_START To lastI
        If wsI.Cells(i, 7).Value <> "" Then
            Dim eanI As String: eanI = CStr(wsI.Cells(i, 2).Value)
            Dim gezaehlt As Double: gezaehlt = Val(wsI.Cells(i, 7).Value)
            For j = 5 To lastA
                If CStr(wsA.Cells(j, colEAN).Value) = eanI Then
                    wsA.Cells(j, colAnz).Value = gezaehlt
                    updated = updated + 1
                    Exit For
                End If
            Next j
        End If
    Next i

    Schnellansicht_Aktualisieren
    MsgBox updated & " Bestaende aktualisiert.", vbInformation
End Sub

' ================================================================
'  INV-SUCHE - SHEET ERSTELLEN
' ================================================================
Sub InvSuche_Setup(Optional silent As Boolean = False)
    Dim wsIS As Worksheet: Set wsIS = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "InvSuche" Then Set wsIS = ws: Exit For
    Next ws
    If wsIS Is Nothing Then
        Set wsIS = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsIS.Name = "InvSuche"
    End If

    Application.ScreenUpdating = False
    wsIS.Cells.Clear
    wsIS.Cells.Interior.ColorIndex = xlNone

    Dim blau As Long: blau = RGB(31, 56, 100)

    ' --- Zeile 1: Titel ---
    wsIS.Range("A1:G1").Merge
    wsIS.Cells(1, 1).Value = "INVENTUR - ARTIKELSUCHE"
    wsIS.Cells(1, 1).Interior.Color = blau
    wsIS.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsIS.Cells(1, 1).Font.Size = 14
    wsIS.Cells(1, 1).Font.Bold = True
    wsIS.Cells(1, 1).HorizontalAlignment = xlCenter
    wsIS.Rows(1).RowHeight = 30

    ' --- Zeile 2: Suchleiste ---
    wsIS.Rows(2).Interior.Color = RGB(235, 243, 250)
    wsIS.Cells(2, 1).Value = "Suche:"
    wsIS.Cells(2, 1).Font.Bold = True
    wsIS.Cells(2, 1).Font.Size = 13
    wsIS.Cells(2, 1).HorizontalAlignment = xlRight
    wsIS.Cells(2, 1).Font.Color = blau
    wsIS.Range("B2:D2").Merge
    wsIS.Cells(2, 2).Interior.Color = RGB(255, 255, 255)
    wsIS.Cells(2, 2).Font.Size = 13
    wsIS.Cells(2, 2).NumberFormat = "@"
    wsIS.Cells(2, 5).Value = "SUCHEN"
    wsIS.Cells(2, 5).Interior.Color = RGB(46, 134, 193)
    wsIS.Cells(2, 5).Font.Color = RGB(255, 255, 255)
    wsIS.Cells(2, 5).Font.Bold = True
    wsIS.Cells(2, 5).Font.Size = 12
    wsIS.Cells(2, 5).HorizontalAlignment = xlCenter
    wsIS.Cells(2, 6).Value = "LEEREN"
    wsIS.Cells(2, 6).Interior.Color = RGB(130, 130, 130)
    wsIS.Cells(2, 6).Font.Color = RGB(255, 255, 255)
    wsIS.Cells(2, 6).Font.Bold = True
    wsIS.Cells(2, 6).Font.Size = 12
    wsIS.Cells(2, 6).HorizontalAlignment = xlCenter
    wsIS.Cells(2, 7).Font.Color = RGB(80, 80, 80)
    wsIS.Cells(2, 7).Font.Italic = True
    wsIS.Rows(2).RowHeight = 32
    wsIS.Range("A2:G2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    wsIS.Range("A2:G2").Borders(xlEdgeBottom).Weight = xlMedium
    wsIS.Range("A2:G2").Borders(xlEdgeBottom).Color = blau

    ' --- Zeile 3: Spaltenkoepfe mit unterschiedlichen Grautönen ---
    Dim hdrs As Variant
    hdrs = Array("Nr", "ArtNr", "Artikel", "SOLL", "VK-Preis", "EAN", "Lagerort")
    ' Exakte Farben aus Lager_Wunsch: B0B0B0, C8C8C8, 9E9E9E, C4C4C4, B4B4B4, D4D4D4, BEBEBE
    Dim grauTone As Variant
    grauTone = Array(RGB(176, 176, 176), RGB(200, 200, 200), RGB(158, 158, 158), _
                     RGB(196, 196, 196), RGB(180, 180, 180), RGB(212, 212, 212), RGB(190, 190, 190))
    Dim j As Integer
    For j = 0 To 6
        wsIS.Cells(3, j + 1).Value = hdrs(j)
        wsIS.Cells(3, j + 1).Interior.Color = grauTone(j)
        wsIS.Cells(3, j + 1).Font.Color = RGB(40, 40, 40)
        wsIS.Cells(3, j + 1).Font.Bold = True
        wsIS.Cells(3, j + 1).Font.Size = 12
        wsIS.Cells(3, j + 1).HorizontalAlignment = xlCenter
        wsIS.Cells(3, j + 1).VerticalAlignment = xlCenter
    Next j
    With wsIS.Range("A3:G3").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(110, 110, 110)
    End With
    wsIS.Range("A3:G3").BorderAround xlContinuous, xlMedium
    wsIS.Rows(3).RowHeight = 26

    ' --- Spaltenbreiten (exakt aus Lager_Wunsch) ---
    wsIS.Columns(1).ColumnWidth = 15
    wsIS.Columns(2).ColumnWidth = 26
    wsIS.Columns(3).ColumnWidth = 43
    wsIS.Columns(4).ColumnWidth = 9
    wsIS.Columns(5).ColumnWidth = 13
    wsIS.Columns(6).ColumnWidth = 17
    wsIS.Columns(7).ColumnWidth = 15
    wsIS.Columns(8).Hidden = True      ' H verborgen, speichert Artikel-Zeilennummer
    wsIS.Columns(2).NumberFormat = "@"
    wsIS.Columns(6).NumberFormat = "@"

    ' --- Freeze: nur Zeilen 1-3 ---
    wsIS.Activate
    wsIS.Cells(4, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 3
    ActiveWindow.SplitColumn = 0
    ActiveWindow.FreezePanes = True

    ' --- Sheet-Events installieren ---
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsIS.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim c As String: c = ""
    c = c & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    c = c & "    If Target.Address = ""$B$2"" Then" & Chr(10)
    c = c & "        On Error GoTo Fehler" & Chr(10)
    c = c & "        Application.EnableEvents = False" & Chr(10)
    c = c & "        LagerMakros.InvSuche_Suchen" & Chr(10)
    c = c & "        Application.EnableEvents = True" & Chr(10)
    c = c & "        Exit Sub" & Chr(10)
    c = c & "Fehler: Application.EnableEvents = True" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10)
    c = c & "    On Error Resume Next" & Chr(10)
    c = c & "    If Target.Row >= 4 And Target.Column >= 1 And Target.Column <= 7 Then" & Chr(10)
    c = c & "        If Me.Cells(Target.Row, 3).Value <> """" Then LagerMakros.InvSuche_ArtikelWaehlen Target.Row" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    Cancel = True" & Chr(10)
    c = c & "    If Target.Row = 2 And Target.Column = 5 Then LagerMakros.InvSuche_Suchen" & Chr(10)
    c = c & "    If Target.Row = 2 And Target.Column = 6 Then LagerMakros.InvSuche_FilterLoeschen" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c

    ' --- UserForm installieren ---
    InvSuche_Form_Installieren

    Application.ScreenUpdating = True
    If Not silent Then MsgBox "InvSuche-Sheet erstellt!", vbInformation
End Sub

' ================================================================
'  INV-SUCHE - USERFORM INSTALLIEREN
' ================================================================
Sub InvSuche_Form_Installieren()
    Dim vbp As Object: Set vbp = ThisWorkbook.VBProject

    ' --- Alte Form loeschen ---
    Dim comp As Object
    Dim tryAgain As Boolean: tryAgain = True
    Do While tryAgain
        tryAgain = False
        On Error Resume Next
        For Each comp In vbp.VBComponents
            If comp.Name = "frmInvSuche" Then
                vbp.VBComponents.Remove comp: tryAgain = True: Exit For
            End If
        Next comp
        On Error GoTo 0
    Loop
    DoEvents

    ' --- Neue Form: NUR Shell, KEIN Designer-Zugriff ---
    Dim frm As Object
    On Error GoTo FormFehler
    Set frm = vbp.VBComponents.Add(3)
    frm.Name = "frmInvSuche"
    DoEvents

    ' --- Alles als Code injizieren ---
    ' Controls werden in UserForm_Initialize() zur Laufzeit erzeugt.
    ' Das vermeidet alle frm.Designer-Probleme.
    Dim fc As Object: Set fc = frm.CodeModule
    Dim s As String: s = ""
    Dim n As String: n = Chr(10)    ' Zeilenumbruch-Shortcut

    ' UserForm_Initialize: baut alle Controls zur Laufzeit auf
    s = s & "Private Sub UserForm_Initialize()" & n
    s = s & "    Me.Caption = ""Artikel-Detail""" & n
    s = s & "    Me.Width = 334" & n
    s = s & "    Me.KeyPreview = True" & n
    s = s & "    Dim ct As Object" & n
    s = s & "    Set ct = Me.Controls.Add(""Forms.Label.1"", ""lblName"")" & n
    s = s & "    ct.Left=0:ct.Top=0:ct.Width=328:ct.Height=44" & n
    s = s & "    ct.BackColor=RGB(46,125,50):ct.ForeColor=RGB(255,255,255)" & n
    s = s & "    ct.BackStyle=1:ct.Font.Bold=True:ct.Font.Size=11:ct.WordWrap=True" & n
    s = s & "    Dim y As Integer : y = 46" & n

    ' Felder: label text | control name | bg R,G,B | fg R,G,B
    Dim fL(6) As String, fN(6) As String
    Dim fBR(6) As Integer, fBG2(6) As Integer, fBB(6) As Integer
    Dim fFR(6) As Integer, fFG2(6) As Integer, fFB(6) As Integer
    fL(0) = "ArtNr:": fN(0) = "lblArtNr": fBR(0) = 245: fBG2(0) = 245: fBB(0) = 245: fFR(0) = 60: fFG2(0) = 60: fFB(0) = 60
    fL(1) = "EAN:": fN(1) = "lblEAN": fBR(1) = 235: fBG2(1) = 245: fBB(1) = 255: fFR(1) = 60: fFG2(1) = 60: fFB(1) = 60
    fL(2) = "Lagerort:": fN(2) = "lblLag": fBR(2) = 235: fBG2(2) = 248: fBB(2) = 235: fFR(2) = 60: fFG2(2) = 60: fFB(2) = 60
    fL(3) = "VK-Preis:": fN(3) = "lblVK": fBR(3) = 245: fBG2(3) = 245: fBB(3) = 245: fFR(3) = 60: fFG2(3) = 60: fFB(3) = 60
    fL(4) = "EK-Preis:": fN(4) = "lblEK": fBR(4) = 245: fBG2(4) = 245: fBB(4) = 245: fFR(4) = 60: fFG2(4) = 60: fFB(4) = 60
    fL(5) = "MwSt:": fN(5) = "lblMwst": fBR(5) = 245: fBG2(5) = 245: fBB(5) = 245: fFR(5) = 60: fFG2(5) = 60: fFB(5) = 60
    fL(6) = "SOLL:": fN(6) = "lblSoll": fBR(6) = 255: fBG2(6) = 243: fBB(6) = 224: fFR(6) = 180: fFG2(6) = 90: fFB(6) = 0

    Dim k As Integer
    For k = 0 To 6
        Dim bg As String: bg = "RGB(" & fBR(k) & "," & fBG2(k) & "," & fBB(k) & ")"
        Dim fg As String: fg = "RGB(" & fFR(k) & "," & fFG2(k) & "," & fFB(k) & ")"
        s = s & "    Set ct = Me.Controls.Add(""Forms.Label.1"")" & n
        s = s & "    ct.Caption=""" & fL(k) & """:ct.Left=2:ct.Top=y:ct.Width=90:ct.Height=24" & n
        s = s & "    ct.BackColor=" & bg & ":ct.BackStyle=1:ct.ForeColor=" & fg & ":ct.Font.Bold=True" & n
        s = s & "    Set ct = Me.Controls.Add(""Forms.Label.1"", """ & fN(k) & """)" & n
        s = s & "    ct.Left=94:ct.Top=y:ct.Width=230:ct.Height=24" & n
        s = s & "    ct.BackColor=" & bg & ":ct.BackStyle=1:ct.ForeColor=" & fg & ":ct.Font.Bold=True" & n
        s = s & "    y = y + 26" & n
    Next k

    ' Gezaehlt-Zeile
    s = s & "    Set ct = Me.Controls.Add(""Forms.Label.1"")" & n
    s = s & "    ct.Caption=""Gezaehlt:"":ct.Left=2:ct.Top=y:ct.Width=90:ct.Height=26" & n
    s = s & "    ct.BackColor=RGB(220,248,220):ct.BackStyle=1:ct.ForeColor=RGB(30,100,30):ct.Font.Bold=True" & n
    s = s & "    Set ct = Me.Controls.Add(""Forms.TextBox.1"", ""txtGezaehlt"")" & n
    s = s & "    ct.Left=94:ct.Top=y+2:ct.Width=228:ct.Height=22" & n
    s = s & "    ct.BackColor=RGB(220,248,220):ct.ForeColor=RGB(30,100,30)" & n
    s = s & "    ct.Font.Size=12:ct.Font.Bold=True" & n
    s = s & "    y = y + 34" & n

    ' Buttons
    s = s & "    y = y + 6" & n
    s = s & "    Set ct = Me.Controls.Add(""Forms.CommandButton.1"", ""btnEintragen"")" & n
    s = s & "    ct.Caption=""EINTRAGEN"":ct.Left=4:ct.Top=y:ct.Width=200:ct.Height=32" & n
    s = s & "    ct.BackColor=RGB(180,90,0):ct.ForeColor=RGB(255,255,255):ct.Font.Bold=True:ct.Font.Size=12" & n
    s = s & "    Set ct = Me.Controls.Add(""Forms.CommandButton.1"", ""btnSchliessen"")" & n
    s = s & "    ct.Caption=""X"":ct.Left=210:ct.Top=y:ct.Width=112:ct.Height=32" & n
    s = s & "    ct.BackColor=RGB(130,130,130):ct.ForeColor=RGB(255,255,255):ct.Font.Bold=True:ct.Font.Size=12" & n
    s = s & "    Me.Height = y + 66" & n
    s = s & "End Sub" & n

    ' Init-Prozedur
    s = s & "Public Sub Init(n As String, nr As String, e As String, l As String, v As String, ek As String, m As String, so As String)" & n
    s = s & "    Me.Controls(""lblName"").Caption = n" & n
    s = s & "    Me.Controls(""lblArtNr"").Caption = nr" & n
    s = s & "    Me.Controls(""lblEAN"").Caption = e" & n
    s = s & "    Me.Controls(""lblLag"").Caption = l" & n
    s = s & "    Me.Controls(""lblVK"").Caption = v" & n
    s = s & "    Me.Controls(""lblEK"").Caption = ek" & n
    s = s & "    Me.Controls(""lblMwst"").Caption = m" & n
    s = s & "    Me.Controls(""lblSoll"").Caption = so" & n
    s = s & "    Me.Controls(""txtGezaehlt"").Value = """"" & n
    s = s & "    Me.Controls(""txtGezaehlt"").SetFocus" & n
    s = s & "End Sub" & n

    ' Tastatur: Enter = Eintragen, Esc = Schliessen
    s = s & "Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & n
    s = s & "    If KeyCode = 13 Then btnEintragen_Click" & n
    s = s & "    If KeyCode = 27 Then Unload Me" & n
    s = s & "End Sub" & n

    ' Button-Events (MSForms loest per Name auf)
    s = s & "Private Sub btnEintragen_Click()" & n
    s = s & "    Dim v As String : v = Trim(Me.Controls(""txtGezaehlt"").Value)" & n
    s = s & "    If v = """" Then MsgBox ""Bitte Menge eingeben."", vbExclamation : Exit Sub" & n
    s = s & "    LagerMakros.InvSuche_Eintragen Me.Controls(""lblName"").Caption, v" & n
    s = s & "    Unload Me" & n
    s = s & "End Sub" & n
    s = s & "Private Sub btnSchliessen_Click()" & n
    s = s & "    Unload Me" & n
    s = s & "End Sub" & n

    fc.AddFromString s
    Exit Sub
FormFehler:
    MsgBox "Fehler beim Erstellen der UserForm:" & Chr(10) & Err.Description, vbCritical, "Form-Fehler"
End Sub

' ================================================================
'  INV-SUCHE - SUCHEN
' ================================================================
Sub InvSuche_Suchen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsIS As Worksheet: Set wsIS = GetSheet("InvSuche")
    If wsA Is Nothing Or wsIS Is Nothing Then Exit Sub

    Dim such As String: such = Trim(wsIS.Cells(2, 2).Value)

    Dim lastRow As Long: lastRow = wsIS.Cells(wsIS.Rows.Count, 3).End(xlUp).Row
    If lastRow >= 4 Then
        wsIS.Range("A4:H" & lastRow).ClearContents
        Dim rr As Long
        For rr = 4 To lastRow
            wsIS.Range("A" & rr & ":G" & rr).Interior.ColorIndex = xlNone
            wsIS.Range("A" & rr & ":G" & rr).Borders.LineStyle = xlNone
        Next rr
    End If

    Application.ScreenUpdating = False

    If such = "" Then
        wsIS.Cells(2, 7).Value = ""
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim colVK  As Long: colVK = Spalte_Finden(wsA, "VK-PREIS")
    Dim colNr  As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    If colArt = 0 Then Application.ScreenUpdating = True: Exit Sub

    Dim woerter() As String: woerter = Split(LCase(such), " ")
    Dim nurZahlen As Boolean: nurZahlen = (such = CStr(Val(such)) And Val(such) > 0)
    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
    Dim sRow As Long: sRow = 4
    Dim nr As Long: nr = 1
    Dim hellgrauS As Long: hellgrauS = RGB(242, 242, 242)
    Dim i As Long, w As Integer, passt As Boolean, suchIn As String

    For i = 5 To lastA
        If wsA.Cells(i, colArt).Value <> "" Then
            If nurZahlen Then
                suchIn = LCase(wsA.Cells(i, colArt).Value & " " & wsA.Cells(i, colEAN).Value)
            Else
                suchIn = LCase(wsA.Cells(i, colArt).Value)
            End If
            passt = True
            For w = 0 To UBound(woerter)
                If Trim(woerter(w)) <> "" Then
                    If InStr(suchIn, Trim(woerter(w))) = 0 Then passt = False: Exit For
                End If
            Next w
            If passt Then
                wsIS.Cells(sRow, 1).Value = nr
                If colNr > 0 Then wsIS.Cells(sRow, 2).Value = CStr(wsA.Cells(i, colNr).Value)
                wsIS.Cells(sRow, 3).Value = wsA.Cells(i, colArt).Value
                If colAnz > 0 Then wsIS.Cells(sRow, 4).Value = wsA.Cells(i, colAnz).Value
                If colVK > 0 Then
                    wsIS.Cells(sRow, 5).Value = wsA.Cells(i, colVK).Value
                    wsIS.Cells(sRow, 5).NumberFormat = "0.00"
                End If
                If colEAN > 0 Then wsIS.Cells(sRow, 6).Value = CStr(wsA.Cells(i, colEAN).Value)
                If colLag > 0 Then wsIS.Cells(sRow, 7).Value = wsA.Cells(i, colLag).Value
                wsIS.Cells(sRow, 8).Value = i
                If sRow Mod 2 = 0 Then wsIS.Range("A" & sRow & ":G" & sRow).Interior.Color = hellgrauS
                ' Vollstaendige Tabellenrahmen
                With wsIS.Range("A" & sRow & ":G" & sRow).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(150, 150, 150)
                End With
                wsIS.Range("A" & sRow & ":G" & sRow).Borders(xlEdgeBottom).Weight = xlMedium
                ' Schrift 12, zentriert
                Dim c As Integer
                For c = 1 To 7
                    wsIS.Cells(sRow, c).Font.Size = 12
                    wsIS.Cells(sRow, c).Font.Bold = False
                    wsIS.Cells(sRow, c).HorizontalAlignment = xlCenter
                    wsIS.Cells(sRow, c).VerticalAlignment = xlCenter
                Next c
                wsIS.Cells(sRow, 3).HorizontalAlignment = xlLeft
                wsIS.Rows(sRow).RowHeight = 21.95
                sRow = sRow + 1
                nr = nr + 1
            End If
        End If
    Next i

    wsIS.Cells(2, 7).Value = (nr - 1) & " Treffer"
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  INV-SUCHE - FILTER LEEREN
' ================================================================
Sub InvSuche_FilterLoeschen()
    On Error GoTo Fehler
    Dim wsIS As Worksheet: Set wsIS = GetSheet("InvSuche")
    If wsIS Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim lastRow As Long: lastRow = wsIS.Cells(wsIS.Rows.Count, 3).End(xlUp).Row
    If lastRow >= 4 Then
        wsIS.Range("A4:H" & lastRow).ClearContents
        Dim rr As Long
        For rr = 4 To lastRow
            wsIS.Range("A" & rr & ":G" & rr).Interior.ColorIndex = xlNone
            wsIS.Range("A" & rr & ":G" & rr).Borders.LineStyle = xlNone
        Next rr
    End If
    wsIS.Cells(2, 2).Value = ""
    wsIS.Cells(2, 7).Value = ""
    g_InvSucheArtikelZeile = 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    wsIS.Cells(2, 2).Select
    Exit Sub
Fehler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  INV-SUCHE - ARTIKEL WAEHLEN (Klick auf Listenzeile -> Popup)
' ================================================================
Sub InvSuche_ArtikelWaehlen(listZeile As Long)
    On Error GoTo Fehler
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsIS As Worksheet: Set wsIS = GetSheet("InvSuche")
    If wsA Is Nothing Or wsIS Is Nothing Then Exit Sub
    If Trim(CStr(wsIS.Cells(listZeile, 3).Value)) = "" Then Exit Sub

    g_InvSucheArtikelZeile = Val(wsIS.Cells(listZeile, 8).Value)
    If g_InvSucheArtikelZeile = 0 Then Exit Sub

    Dim colEAN  As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt  As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colAnz  As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim colVK   As Long: colVK = Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK   As Long: colEK = Spalte_Finden(wsA, "EK-PREIS")
    Dim colNr   As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colLag  As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colMwst As Long: colMwst = Spalte_Finden(wsA, "MWST")
    Dim az As Long: az = g_InvSucheArtikelZeile

    ' Zeile hervorheben
    Dim lastRow As Long: lastRow = wsIS.Cells(wsIS.Rows.Count, 3).End(xlUp).Row
    Dim r As Long
    For r = 4 To lastRow
        If r = listZeile Then
            wsIS.Range("A" & r & ":G" & r).Interior.Color = RGB(184, 204, 228)
        ElseIf r Mod 2 = 0 Then
            wsIS.Range("A" & r & ":G" & r).Interior.Color = RGB(242, 242, 242)
        Else
            wsIS.Range("A" & r & ":G" & r).Interior.ColorIndex = xlNone
        End If
    Next r

    ' Werte fuer Popup zusammenstellen
    Dim artName As String: artName = wsA.Cells(az, colArt).Value
    Dim artNr   As String: artNr = IIf(colNr > 0, CStr(wsA.Cells(az, colNr).Value), "")
    Dim ean     As String: ean = IIf(colEAN > 0, CStr(wsA.Cells(az, colEAN).Value), "")
    Dim lagort  As String: lagort = IIf(colLag > 0, wsA.Cells(az, colLag).Value, "")
    Dim vk      As String: vk = IIf(colVK > 0, Format(wsA.Cells(az, colVK).Value, "0.00") & " EUR", "")
    Dim ek      As String: ek = IIf(colEK > 0, Format(wsA.Cells(az, colEK).Value, "0.00") & " EUR", "")
    Dim mwst    As String: mwst = IIf(colMwst > 0, wsA.Cells(az, colMwst).Value & " %", "19 %")
    Dim soll    As String: soll = IIf(colAnz > 0, CStr(wsA.Cells(az, colAnz).Value), "0")

    ' Popup anzeigen
    Dim frm As Object
    Set frm = Nothing
    On Error Resume Next
    Set frm = vbA.UserForms.Add("frmInvSuche")
    On Error GoTo Fehler
    If frm Is Nothing Then
        MsgBox "UserForm fehlt. Bitte Setup_Ausfuehren nochmal starten.", vbExclamation
        Exit Sub
    End If
    frm.Init artName, artNr, ean, lagort, vk, ek, mwst, soll
    frm.Show vbModal
    Exit Sub
Fehler:
    Application.EnableEvents = True
End Sub

' ================================================================
'  INV-SUCHE - EINTRAGEN (wird vom Popup aufgerufen)
' ================================================================
Sub InvSuche_Eintragen(artName As String, mengeStr As String)
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Then Exit Sub
    If g_InvSucheArtikelZeile = 0 Then Exit Sub

    Dim menge As Double: menge = Val(mengeStr)
    Dim az As Long: az = g_InvSucheArtikelZeile

    If Not wsI Is Nothing Then
        Dim colEAN_I As Long: colEAN_I = Spalte_Finden(wsA, "EAN13")
        Dim colArt_I As Long: colArt_I = Spalte_Finden(wsA, "ARTIKEL")
        Dim colLag_I As Long: colLag_I = Spalte_Finden(wsA, "LAGERORT")
        Dim colEK_I  As Long: colEK_I = Spalte_Finden(wsA, "EK-PREIS")
        Dim colAnz_I As Long: colAnz_I = Spalte_Finden(wsA, "ANZAHL")
        Dim lastInv As Long: lastInv = wsI.Cells(wsI.Rows.Count, 3).End(xlUp).Row
        Dim gefunden As Boolean: gefunden = False
        Dim ii As Long
        For ii = INV_DATEN_START To lastInv
            If LCase(Trim(wsI.Cells(ii, 3).Value)) = LCase(artName) Then
                wsI.Cells(ii, 7).Value = menge: gefunden = True: Exit For
            End If
        Next ii
        If Not gefunden Then
            Dim nRow As Long
            If lastInv < INV_DATEN_START Then nRow = INV_DATEN_START Else nRow = lastInv + 1
            wsI.Cells(nRow, 1).Value = nRow - INV_DATEN_START + 1
            If colEAN_I > 0 Then wsI.Cells(nRow, 2).Value = wsA.Cells(az, colEAN_I).Value
            wsI.Cells(nRow, 3).Value = wsA.Cells(az, colArt_I).Value
            If colLag_I > 0 Then wsI.Cells(nRow, 4).Value = wsA.Cells(az, colLag_I).Value
            If colEK_I > 0 Then wsI.Cells(nRow, 5).Value = wsA.Cells(az, colEK_I).Value: wsI.Cells(nRow, 5).NumberFormat = "0.00"
            If colAnz_I > 0 Then wsI.Cells(nRow, 6).Value = wsA.Cells(az, colAnz_I).Value
            wsI.Cells(nRow, 7).Value = menge
            wsI.Cells(nRow, 8).Formula = "=IF(G" & nRow & "="""","""",G" & nRow & "-F" & nRow & ")"
            wsI.Cells(nRow, 9).Formula = "=IF(G" & nRow & "<>"""",E" & nRow & "*G" & nRow & ",E" & nRow & "*F" & nRow & ")"
            wsI.Cells(nRow, 9).NumberFormat = "0.00"
        End If
    End If

    MsgBox artName & Chr(10) & "Gezaehlt: " & Format(menge, "0") & " Stk  ->  eingetragen", vbInformation, "Gespeichert"
    g_InvSucheArtikelZeile = 0
End Sub

' ================================================================
'  SPALTEN ANZEIGEN (einmalig zur Diagnose)
' ================================================================
Sub Spalten_Anzeigen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then MsgBox "Artikel-Sheet nicht gefunden!", vbCritical: Exit Sub
    Dim s As String: s = "Spalten in Zeile 2:" & Chr(10) & Chr(10)
    Dim i As Integer
    For i = 1 To 30
        If wsA.Cells(2, i).Value <> "" Then
            s = s & "Spalte " & i & ": " & wsA.Cells(2, i).Value & Chr(10)
        End If
    Next i
    MsgBox s, vbInformation, "Artikel-Spalten"
End Sub

