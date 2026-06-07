Attribute VB_Name = "NeuArtikelModul"
Option Explicit

' ================================================================
'  NEUER ARTIKEL + SUCHE & SCANNER
'  Fuer: 2026_Lagerverwaltung_V2.xlsm
'  Einmal AllesNeuEinrichten ausfuehren - fertig!
' ================================================================

Sub Setup_ArtikelKopf()
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsA Is Nothing Then MsgBox "Artikel-Blatt nicht gefunden!", vbCritical: Exit Sub

    Application.ScreenUpdating = False

    ' --- Zeile 1 vorbereiten ---
    wsA.Rows(1).RowHeight = 26
    wsA.Rows(1).UnMerge

    ' Alte Shapes in Zeile 1 loeschen
    Dim shp As Shape
    For Each shp In wsA.Shapes
        If shp.Top < wsA.Rows(1).Height + 2 Then shp.Delete
    Next shp

    ' Zellen leeren (Zeile 1)
    wsA.Rows(1).ClearContents
    wsA.Rows(1).Interior.Color = RGB(242, 242, 242)

    ' --- Suchfeld (Zelle A1:D1) ---
    wsA.Range("A1:D1").Merge
    wsA.Range("A1").Value = ""
    wsA.Range("A1").Interior.Color = RGB(255, 255, 255)
    wsA.Range("A1").Font.Size = 11
    wsA.Range("A1").HorizontalAlignment = xlLeft
    wsA.Range("A1").IndentLevel = 1
    With wsA.Range("A1").Borders
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With

    ' Platzhalter-Text (wird beim Tippen ueberschrieben)
    wsA.Range("A1").Value = "  Suche..."
    wsA.Range("A1").Font.Color = RGB(150, 150, 150)
    wsA.Range("A1").Font.Italic = True

    ' --- Shapes: Button-Definitionen ---
    ' Format: Name, Beschriftung, Farbe, Makro
    Dim btnData(4, 3) As String
    btnData(0, 0) = "btn_NeuerArtikel":  btnData(0, 1) = "Neuer Artikel":  btnData(0, 2) = "16737945": btnData(0, 3) = "NeuArtikelModul.NeuerArtikel_Oeffnen"
    btnData(1, 0) = "btn_ZuAbgang":      btnData(1, 1) = "Zu-/Abgang":     btnData(1, 2) = "4561153":  btnData(1, 3) = "NeuArtikelModul.ZuAbgang_Oeffnen"
    btnData(2, 0) = "btn_Etikett":       btnData(2, 1) = "Etikett drucken": btnData(2, 2) = "2039583": btnData(2, 3) = "NeuArtikelModul.Etikett_Drucken"
    btnData(3, 0) = "btn_EKToggle":      btnData(3, 1) = "EK ausblenden":  btnData(3, 2) = "12611584": btnData(3, 3) = "NeuArtikelModul.EK_Toggle"
    btnData(4, 0) = "btn_Schnell":       btnData(4, 1) = "Schnellansicht": btnData(4, 2) = "2920428":  btnData(4, 3) = "NeuArtikelModul.Schnellansicht_Oeffnen"

    ' Position: nach Spalte E starten
    Dim startLeft As Double: startLeft = wsA.Columns("E").Left + 4
    Dim btnTop    As Double: btnTop = 3
    Dim btnH      As Double: btnH = 20
    Dim btnW      As Double: btnW = 95
    Dim gap       As Double: gap = 4
    Dim i As Integer

    For i = 0 To 4
        ' Alten Button gleichen Namens loeschen
        On Error Resume Next
        wsA.Shapes(btnData(i, 0)).Delete
        On Error GoTo 0

        Dim s As Shape
        Set s = wsA.Shapes.AddShape(msoShapeRoundedRectangle, _
            startLeft + i * (btnW + gap), btnTop, btnW, btnH)
        s.Name = btnData(i, 0)
        s.TextFrame.Characters.Text = btnData(i, 1)
        s.TextFrame.Characters.Font.Bold = True
        s.TextFrame.Characters.Font.Size = 9
        s.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        s.TextFrame.HorizontalAlignment = xlHAlignCenter
        s.TextFrame.VerticalAlignment = xlVAlignCenter
        s.Fill.ForeColor.RGB = CLng(btnData(i, 2))
        s.Line.Visible = msoFalse
        s.OnAction = btnData(i, 3)
    Next i

    ' --- Filter-X Button (kleiner, neben Suchfeld) ---
    On Error Resume Next
    wsA.Shapes("btn_FilterX").Delete
    On Error GoTo 0
    Dim fx As Shape
    Set fx = wsA.Shapes.AddShape(msoShapeRoundedRectangle, _
        wsA.Range("E1").Left - 26, btnTop, 22, btnH)
    fx.Name = "btn_FilterX"
    fx.TextFrame.Characters.Text = "X"
    fx.TextFrame.Characters.Font.Bold = True
    fx.TextFrame.Characters.Font.Size = 9
    fx.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    fx.Fill.ForeColor.RGB = RGB(192, 0, 0)
    fx.Line.Visible = msoFalse
    fx.OnAction = "NeuArtikelModul.Artikel_FilterLoeschen"

    Application.ScreenUpdating = True
    MsgBox "Artikel-Kopf fertig! Suchfeld in A1, alle Buttons gesetzt.", vbInformation
End Sub


' ================================================================
'  Platzhalter-Funktionen (werden spaeter befuellt)
' ================================================================
Sub ZuAbgang_Oeffnen()
    MsgBox "Zu-/Abgang kommt noch!", vbInformation
End Sub
Sub Etikett_Drucken()
    MsgBox "Etikett kommt noch!", vbInformation
End Sub
Sub EK_Toggle()
    LagerMakros.EK_Ausblenden
End Sub
Sub Schnellansicht_Oeffnen()
    ThisWorkbook.Sheets("Schnellansicht").Activate
End Sub
Sub Artikel_FilterLoeschen()
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    On Error Resume Next
    wsA.Rows("3:50000").Hidden = False
    wsA.ShowAllData
    wsA.Shapes("lbl_Treffer").Delete
    wsA.Range("A1").Value = "  Suche..."
    wsA.Range("A1").Font.Color = RGB(150, 150, 150)
    wsA.Range("A1").Font.Italic = True
    On Error GoTo 0
End Sub

' ================================================================
'  Artikel-Sheet: Suche mit Row-Hiding + Trefferanzeige
' ================================================================
Sub Artikel_Suchen(wsA As Worksheet, such As String)
    ' Alle Zeilen einblenden BEVOR lastRow berechnet wird
    ' (End(xlUp) findet sonst nur sichtbare Zeilen)
    wsA.Rows("3:50000").Hidden = False
    If wsA.AutoFilterMode Then wsA.AutoFilter.ShowAllData

    Dim cArt As Long: cArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim cEAN As Long: cEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    If cArt = 0 Then Exit Sub

    Dim lastRow As Long: lastRow = wsA.Cells(wsA.Rows.count, cArt).End(xlUp).Row

    ' Zahl >= 8 Stellen -> EAN-Suche, sonst Artikelname mit AND-Mehrwortsuche
    Dim nurEAN As Boolean: nurEAN = (Len(such) >= 8 And IsNumeric(such))
    Dim woerter() As String: woerter = Split(LCase(such), " ")

    Dim i As Long, w As Integer, passt As Boolean, suchIn As String
    Dim treffer As Long: treffer = 0

    Application.ScreenUpdating = False
    For i = 3 To lastRow
        If nurEAN Then
            suchIn = LCase(CStr(wsA.Cells(i, cEAN).Value))
        Else
            suchIn = LCase(CStr(wsA.Cells(i, cArt).Value))
        End If
        passt = (suchIn <> "")
        If passt Then
            For w = 0 To UBound(woerter)
                If Trim(woerter(w)) <> "" Then
                    If InStr(suchIn, Trim(woerter(w))) = 0 Then
                        passt = False: Exit For
                    End If
                End If
            Next w
        End If
        wsA.Rows(i).Hidden = Not passt
        If passt Then treffer = treffer + 1
    Next i

    ' Shape lbl_Treffer: altes loeschen, neues anzeigen
    On Error Resume Next
    wsA.Shapes("lbl_Treffer").Delete
    On Error GoTo 0
    Dim shp As Shape
    Set shp = wsA.Shapes.AddShape(msoShapeRoundedRectangle, _
        wsA.Range("A1").Left + 2, wsA.Range("A1").Top + 3, 75, 19)
    shp.Name = "lbl_Treffer"
    shp.TextFrame.Characters.Text = treffer & " Treffer"
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = IIf(treffer = 0, RGB(192, 0, 0), RGB(46, 125, 50))
    shp.Line.Visible = msoFalse
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.Placement = xlFreeFloating

    Application.ScreenUpdating = True
End Sub


Sub AllesNeuEinrichten()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo Fehler
    Setup_SucheScanner
    Setup_NeuerArtikel
    DropdownsEinrichten

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Fertig!" & Chr(10) & Chr(10) & _
           "Zwei neue Blaetter erstellt:" & Chr(10) & _
           "  - Suche & Scanner" & Chr(10) & _
           "  - Neuer Artikel" & Chr(10) & Chr(10) & _
           "Zum Testen: Blatt 'Neuer Artikel' oeffnen.", _
           vbInformation, "Setup abgeschlossen"
    Exit Sub
Fehler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, vbCritical
End Sub


' ================================================================
'  SETUP - Suche & Scanner Sheet
' ================================================================
Sub Setup_SucheScanner()
    ' Sheet loeschen falls vorhanden
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Suche & Scanner").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim wsSS As Worksheet
    Set wsSS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsSS.Name = "Suche & Scanner"

    Dim blau As Long: blau = RGB(32, 55, 100)

    ' Zeile 1: Titel + Suchfeld + Buttons
    wsSS.Rows(1).RowHeight = 28
    wsSS.Range("A1:A1").Merge
    wsSS.Cells(1, 1).Value = "Suche & Scanner"
    wsSS.Cells(1, 1).Interior.Color = blau
    wsSS.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsSS.Cells(1, 1).Font.Bold = True
    wsSS.Cells(1, 1).Font.Size = 12

    ' Suchfeld B1
    wsSS.Cells(1, 2).Interior.Color = RGB(255, 255, 153)
    wsSS.Cells(1, 2).Font.Size = 12
    wsSS.Cells(1, 2).Value = ""

    ' Button: FILTER LOESCHEN (C1)
    wsSS.Cells(1, 3).Value = "FILTER LOESCHEN"
    wsSS.Cells(1, 3).Interior.Color = RGB(237, 125, 49)
    wsSS.Cells(1, 3).Font.Color = RGB(255, 255, 255)
    wsSS.Cells(1, 3).Font.Bold = True
    wsSS.Cells(1, 3).HorizontalAlignment = xlCenter

    ' Zeile 2: Spaltenkoepfe
    wsSS.Rows(2).RowHeight = 20
    Dim hdrs As Variant
    hdrs = Array("EAN13", "ARTIKEL", "VK-PREIS", "EK-PREIS", "MWST", "ANZAHL", "EINHEIT", "LAGERORT", "WARENGRUPPE")
    Dim j As Integer
    For j = 0 To UBound(hdrs)
        wsSS.Cells(2, j + 1).Value = hdrs(j)
        wsSS.Cells(2, j + 1).Interior.Color = RGB(46, 80, 144)
        wsSS.Cells(2, j + 1).Font.Color = RGB(255, 255, 255)
        wsSS.Cells(2, j + 1).Font.Bold = True
        wsSS.Cells(2, j + 1).HorizontalAlignment = xlCenter
    Next j

    ' Spaltenbreiten
    wsSS.Columns(1).ColumnWidth = 15  ' EAN
    wsSS.Columns(2).ColumnWidth = 35  ' Artikel
    wsSS.Columns(3).ColumnWidth = 10  ' VK
    wsSS.Columns(4).ColumnWidth = 10  ' EK
    wsSS.Columns(5).ColumnWidth = 7   ' MWST
    wsSS.Columns(6).ColumnWidth = 9   ' Anzahl
    wsSS.Columns(7).ColumnWidth = 8   ' Einheit
    wsSS.Columns(8).ColumnWidth = 14  ' Lagerort
    wsSS.Columns(9).ColumnWidth = 18  ' Warengruppe

    ' Daten aus Artikel-Sheet kopieren
    SucheScanner_Aktualisieren

    ' Events installieren
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsSS.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim c As String: c = ""
    c = c & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    c = c & "    If Target.Address = ""$B$1"" Then" & Chr(10)
    c = c & "        Application.EnableEvents = False" & Chr(10)
    c = c & "        NeuArtikelModul.SucheScanner_Filtern CStr(Target.Value)" & Chr(10)
    c = c & "        Application.EnableEvents = True" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10)
    c = c & "    If Target.Address = ""$C$1"" Then" & Chr(10)
    c = c & "        Application.EnableEvents = False" & Chr(10)
    c = c & "        Me.Range(""B1"").Value = """"" & Chr(10)
    c = c & "        NeuArtikelModul.SucheScanner_Filtern """"" & Chr(10)
    c = c & "        Me.Range(""B1"").Select" & Chr(10)
    c = c & "        Application.EnableEvents = True" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    If Target.Row >= 3 And Target.Column >= 1 Then" & Chr(10)
    c = c & "        Cancel = True" & Chr(10)
    c = c & "        NeuArtikelModul.SucheScanner_ArtikelLaden Target.Row" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_Activate()" & Chr(10)
    c = c & "    Me.Cells(1, 2).Select" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c
    On Error GoTo 0
End Sub


' ================================================================
'  Suche & Scanner - Artikel-Daten laden/aktualisieren
' ================================================================
Sub SucheScanner_Aktualisieren()
    Dim wsSS As Worksheet: Set wsSS = LagerMakros.GetSheet("Suche")
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsSS Is Nothing Or wsA Is Nothing Then Exit Sub

    ' Alte Daten loeschen
    Dim lastSS As Long: lastSS = wsSS.Cells(wsSS.Rows.count, 1).End(xlUp).Row
    If lastSS >= 3 Then wsSS.Range("A3:I" & lastSS).Clear

    ' Spalten im Artikel-Sheet finden
    Dim cEAN  As Long: cEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim cArt  As Long: cArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim cVK   As Long: cVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim cEK   As Long: cEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim cMwst As Long: cMwst = LagerMakros.Spalte_Finden(wsA, "MWST")
    Dim cAnz  As Long: cAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim cEinh As Long: cEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim cLag  As Long: cLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim cWG   As Long: cWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")

    ' Alle Zeilen einblenden BEVOR lastA berechnet wird
    ' (End(xlUp) findet sonst nur sichtbare Zeilen -> fehlende Artikel nach Filter)
    If wsA.AutoFilterMode Then wsA.AutoFilter.ShowAllData
    wsA.Rows("3:50000").Hidden = False
    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.count, cArt).End(xlUp).Row
    Dim sRow As Long: sRow = 3
    Dim i As Long

    Application.ScreenUpdating = False
    For i = 3 To lastA
        If wsA.Cells(i, cArt).Value <> "" Then
            ' EAN als Text
            Dim vEAN As String
            If cEAN > 0 Then
                Dim rawEAN As Variant: rawEAN = wsA.Cells(i, cEAN).Value
                If IsNumeric(rawEAN) And rawEAN <> "" Then
                    vEAN = Format(CDbl(rawEAN), "0000000000000")
                Else
                    vEAN = CStr(rawEAN)
                End If
            End If

            wsSS.Cells(sRow, 1).NumberFormat = "@"
            wsSS.Cells(sRow, 1).Value = vEAN
            wsSS.Cells(sRow, 2).Value = IIf(cArt > 0, wsA.Cells(i, cArt).Value, "")
            wsSS.Cells(sRow, 3).Value = IIf(cVK > 0, wsA.Cells(i, cVK).Value, "")
            wsSS.Cells(sRow, 3).NumberFormat = "#,##0.00"
            wsSS.Cells(sRow, 4).Value = IIf(cEK > 0, wsA.Cells(i, cEK).Value, "")
            wsSS.Cells(sRow, 4).NumberFormat = "#,##0.00"
            wsSS.Cells(sRow, 5).Value = IIf(cMwst > 0, wsA.Cells(i, cMwst).Value, "")
            wsSS.Cells(sRow, 6).Value = IIf(cAnz > 0, wsA.Cells(i, cAnz).Value, "")
            wsSS.Cells(sRow, 7).Value = IIf(cEinh > 0, wsA.Cells(i, cEinh).Value, "")
            wsSS.Cells(sRow, 8).Value = IIf(cLag > 0, wsA.Cells(i, cLag).Value, "")
            wsSS.Cells(sRow, 9).Value = IIf(cWG > 0, wsA.Cells(i, cWG).Value, "")

            ' Bestand 0 = rot markieren
            If IIf(cAnz > 0, val(wsA.Cells(i, cAnz).Value), 1) = 0 Then
                wsSS.Rows(sRow).Interior.Color = RGB(255, 199, 206)
            ElseIf sRow Mod 2 = 0 Then
                wsSS.Rows(sRow).Interior.Color = RGB(242, 242, 242)
            Else
                wsSS.Rows(sRow).Interior.ColorIndex = xlNone
            End If
            sRow = sRow + 1
        End If
    Next i
    Application.ScreenUpdating = True
End Sub


' ================================================================
'  Suche & Scanner - Filtern
' ================================================================
Sub SucheScanner_Filtern(such As String)
    Dim wsSS As Worksheet: Set wsSS = LagerMakros.GetSheet("Suche")
    If wsSS Is Nothing Then Exit Sub

    Dim s As String: s = Trim(LCase(such))
    ' Alle Zeilen einblenden BEVOR lastRow berechnet wird
    ' (sonst werden nach einem vorherigen Filter Zeilen dauerhaft versteckt)
    wsSS.Rows("3:50000").Hidden = False
    Dim lastRow As Long: lastRow = wsSS.Cells(wsSS.Rows.count, 2).End(xlUp).Row
    Dim i As Long

    Application.ScreenUpdating = False
    For i = 3 To lastRow
        If s = "" Then
            wsSS.Rows(i).Hidden = False
        Else
            Dim ean  As String: ean = LCase(CStr(wsSS.Cells(i, 1).Value))
            Dim art  As String: art = LCase(CStr(wsSS.Cells(i, 2).Value))
            Dim lag  As String: lag = LCase(CStr(wsSS.Cells(i, 8).Value))
            Dim wg   As String: wg = LCase(CStr(wsSS.Cells(i, 9).Value))
            wsSS.Rows(i).Hidden = Not (InStr(ean, s) > 0 Or InStr(art, s) > 0 Or _
                                       InStr(lag, s) > 0 Or InStr(wg, s) > 0)
        End If
    Next i
    Application.ScreenUpdating = True
End Sub


' ================================================================
'  Suche & Scanner - Artikel in Neuer Artikel laden (Doppelklick)
' ================================================================
Sub SucheScanner_ArtikelLaden(zeile As Long)
    Dim wsSS As Worksheet: Set wsSS = LagerMakros.GetSheet("Suche")
    Dim wsNA As Worksheet: Set wsNA = LagerMakros.GetSheet("Neuer Artikel")
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsSS Is Nothing Or wsNA Is Nothing Or wsA Is Nothing Then Exit Sub

    Dim sEAN As String: sEAN = CStr(wsSS.Cells(zeile, 1).Value)
    Dim sArt As String: sArt = CStr(wsSS.Cells(zeile, 2).Value)
    If sArt = "" Then Exit Sub

    ' Artikel im Artikel-Sheet suchen
    Dim cEAN  As Long: cEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim cArt  As Long: cArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim cVK   As Long: cVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim cEK   As Long: cEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim cMwst As Long: cMwst = LagerMakros.Spalte_Finden(wsA, "MWST")
    Dim cAnz  As Long: cAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim cEinh As Long: cEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim cLag  As Long: cLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim cWG   As Long: cWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim cNr   As Long: cNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim cAttr As Long: cAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")
    Dim cTA   As Long: cTA = LagerMakros.Spalte_Finden(wsA, "TextA")
    Dim cTB   As Long: cTB = LagerMakros.Spalte_Finden(wsA, "TextB")

    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.count, cArt).End(xlUp).Row
    Dim i As Long

    For i = 3 To lastA
        If CStr(wsA.Cells(i, cArt).Value) = sArt Then
            ' Felder befuellen
            If cNr > 0 Then wsNA.Range("B4").Value = CStr(wsA.Cells(i, cNr).Value)
            wsNA.Range("B5").Value = sArt
            If cEK > 0 Then wsNA.Range("B6").Value = wsA.Cells(i, cEK).Value
            If cVK > 0 Then wsNA.Range("B7").Value = wsA.Cells(i, cVK).Value
            If cWG > 0 Then wsNA.Range("B8").Value = wsA.Cells(i, cWG).Value
            If cLag > 0 Then wsNA.Range("B9").Value = wsA.Cells(i, cLag).Value
            If cMwst > 0 Then wsNA.Range("B10").Value = wsA.Cells(i, cMwst).Value
            If cEinh > 0 Then wsNA.Range("B11").Value = wsA.Cells(i, cEinh).Value
            wsNA.Range("B12").NumberFormat = "@"
            wsNA.Range("B12").Value = sEAN
            If cAttr > 0 Then wsNA.Range("B13").Value = wsA.Cells(i, cAttr).Value
            If cTA > 0 Then wsNA.Range("B14").Value = wsA.Cells(i, cTA).Value
            If cTB > 0 Then wsNA.Range("B15").Value = wsA.Cells(i, cTB).Value
            If cAnz > 0 Then wsNA.Range("B16").Value = wsA.Cells(i, cAnz).Value
            Exit For
        End If
    Next i

    ' Zu Neuer Artikel wechseln
    wsNA.Visible = xlSheetVisible
    wsNA.Activate
End Sub


' ================================================================
'  SETUP - Neuer Artikel Sheet
' ================================================================
Sub Setup_NeuerArtikel()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Neuer Artikel").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim wsNA As Worksheet
    Set wsNA = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsNA.Name = "Neuer Artikel"

    Dim blau   As Long: blau = RGB(32, 55, 100)
    Dim gruen  As Long: gruen = RGB(55, 86, 35)
    Dim gelb   As Long: gelb = RGB(255, 255, 0)
    Dim hellg  As Long: hellg = RGB(255, 255, 204)

    ' Zeile 1: Titel
    wsNA.Range("A1:C1").Merge
    wsNA.Cells(1, 1).Value = "NEUEN ARTIKEL ANLEGEN"
    wsNA.Cells(1, 1).Font.Bold = True: wsNA.Cells(1, 1).Font.Size = 14
    wsNA.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsNA.Cells(1, 1).Interior.Color = blau
    wsNA.Cells(1, 1).HorizontalAlignment = xlCenter
    wsNA.Rows(1).RowHeight = 30

    ' Zeile 2: Hinweis
    wsNA.Range("A2:C2").Merge
    wsNA.Cells(2, 1).Value = "Felder ausfuellen - dann SPEICHERN doppelklicken"
    wsNA.Cells(2, 1).Font.Italic = True
    wsNA.Cells(2, 1).Font.Color = RGB(192, 0, 0)
    wsNA.Cells(2, 1).HorizontalAlignment = xlCenter
    wsNA.Rows(2).RowHeight = 20

    ' Zeile 3: Leerzeile
    wsNA.Rows(3).RowHeight = 8

    ' Felder definieren: Zeile 4-17
    Dim labels As Variant
    labels = Array( _
        "* Artikelnr.:", _
        "Artikel:", _
        "EK-Preis " & ChrW(8364) & ":", _
        "VK-Preis " & ChrW(8364) & ":", _
        "Warengruppe:", _
        "Lagerort:", _
        "MwSt %:", _
        "Einheit:", _
        "EAN13:", _
        "Attribut:", _
        "TextA:", _
        "TextB:", _
        "Anfangsbestand:", _
        "Vermerk:" _
    )

    Dim hints As Variant
    hints = Array( _
        "z.B. CON-MLS-06 (eindeutig!)", _
        "Vollstaendiger Artikelname", _
        "Einkaufspreis netto", _
        "Verkaufspreis brutto", _
        "Dropdown oder tippen -> Vorschlaege rechts", _
        "Dropdown oder tippen -> Vorschlaege rechts", _
        "19 oder 7", _
        "Stk / Pack / kg ...", _
        "Barcode scannen oder tippen", _
        "Tippen -> Vorschlaege rechts", _
        "Freitext A", _
        "Freitext B", _
        "Anfangsmenge beim Anlegen", _
        "Interne Notiz" _
    )

    Dim k As Integer
    For k = 0 To UBound(labels)
        Dim r As Long: r = k + 4
        ' Label
        wsNA.Cells(r, 1).Value = labels(k)
        wsNA.Cells(r, 1).Font.Bold = True
        wsNA.Cells(r, 1).Font.Size = 11
        ' Eingabefeld
        wsNA.Cells(r, 2).Interior.Color = IIf(k Mod 2 = 0, gelb, hellg)
        wsNA.Cells(r, 2).Font.Size = 11
        ' Hinweis
        wsNA.Cells(r, 3).Value = hints(k)
        wsNA.Cells(r, 3).Font.Italic = True
        wsNA.Cells(r, 3).Font.Size = 9
        wsNA.Cells(r, 3).Font.Color = RGB(89, 89, 89)
        wsNA.Rows(r).RowHeight = 22
    Next k

    ' EAN und ArtNr als Text
    wsNA.Range("B4").NumberFormat = "@"
    wsNA.Range("B12").NumberFormat = "@"

    ' Zeile 18: Leerzeile
    wsNA.Rows(18).RowHeight = 10

    ' Zeile 19: Buttons
    wsNA.Cells(19, 1).Value = "SPEICHERN  (Doppelklick!)"
    wsNA.Cells(19, 1).Font.Bold = True: wsNA.Cells(19, 1).Font.Size = 12
    wsNA.Cells(19, 1).Font.Color = RGB(255, 255, 255)
    wsNA.Cells(19, 1).Interior.Color = gruen
    wsNA.Cells(19, 1).HorizontalAlignment = xlCenter
    wsNA.Cells(19, 2).Value = "FELDER LEEREN  (Doppelklick!)"
    wsNA.Cells(19, 2).Font.Bold = True: wsNA.Cells(19, 2).Font.Size = 12
    wsNA.Cells(19, 2).Font.Color = RGB(255, 255, 255)
    wsNA.Cells(19, 2).Interior.Color = RGB(192, 0, 0)
    wsNA.Cells(19, 2).HorizontalAlignment = xlCenter
    wsNA.Rows(19).RowHeight = 28

    ' Spaltenbreiten
    wsNA.Columns(1).ColumnWidth = 18
    wsNA.Columns(2).ColumnWidth = 35
    wsNA.Columns(3).ColumnWidth = 35
    wsNA.Columns(4).ColumnWidth = 25  ' Vorschlaege
    wsNA.Columns(5).ColumnWidth = 1   ' Hilfsspalte Warengruppen
    wsNA.Columns(6).ColumnWidth = 1   ' Hilfsspalte Attribute

    ' Zeile 20: Vorschlags-Ueberschrift
    wsNA.Cells(4, 4).Value = "< Doppelklick auf Vorschlag zum Uebernehmen"
    wsNA.Cells(4, 4).Font.Italic = True
    wsNA.Cells(4, 4).Font.Color = RGB(0, 70, 127)
    wsNA.Cells(4, 4).Font.Size = 9

    ' Blatt verstecken (wie SchnellDetail)
    wsNA.Visible = xlSheetHidden

    ' Events installieren
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsNA.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim c As String: c = ""
    c = c & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    c = c & "    Dim suchZellen As Variant" & Chr(10)
    c = c & "    suchZellen = Array(""$B$8"", ""$B$9"", ""$B$10"", ""$B$13"")" & Chr(10)
    c = c & "    Dim istSuch As Boolean: istSuch = False" & Chr(10)
    c = c & "    Dim v As Variant" & Chr(10)
    c = c & "    For Each v In suchZellen" & Chr(10)
    c = c & "        If Target.Address = v Then istSuch = True" & Chr(10)
    c = c & "    Next v" & Chr(10)
    c = c & "    If Not istSuch Then Exit Sub" & Chr(10)
    c = c & "    Application.EnableEvents = False" & Chr(10)
    c = c & "    NeuArtikelModul.NeuerArtikel_Vorschlaege Target" & Chr(10)
    c = c & "    Application.EnableEvents = True" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    If Target.Address = ""$A$19"" Then" & Chr(10)
    c = c & "        Cancel = True: NeuArtikelModul.NeuerArtikel_Speichern: Exit Sub" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "    If Target.Address = ""$B$19"" Then" & Chr(10)
    c = c & "        Cancel = True: NeuArtikelModul.NeuerArtikel_FelderLeeren: Exit Sub" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "    If Target.Column = 4 And Target.Row >= 5 And Target.Row <= 25 Then" & Chr(10)
    c = c & "        Cancel = True" & Chr(10)
    c = c & "        NeuArtikelModul.NeuerArtikel_VorschlagUebernehmen Target" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c
    On Error GoTo 0
End Sub


' ================================================================
'  Neuer Artikel - Vorschlaege anzeigen beim Tippen
' ================================================================
Sub NeuerArtikel_Vorschlaege(Target As Range)
    Dim wsNA As Worksheet: Set wsNA = Target.Worksheet
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub

    ' Vorschlaege loeschen
    wsNA.Range("D5:D25").ClearContents
    wsNA.Range("D5:D25").Interior.ColorIndex = xlNone

    Dim suchbegriff As String: suchbegriff = Trim(LCase(CStr(Target.Value)))
    If Len(suchbegriff) < 2 Then Exit Sub

    ' Spalte und Titel bestimmen
    Dim suchSpalte As Long
    Dim titel As String
    Select Case Target.Address
        Case "$B$8":  suchSpalte = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE"): titel = "Warengruppe"
        Case "$B$9":  suchSpalte = LagerMakros.Spalte_Finden(wsA, "LAGERORT"):    titel = "Lagerort"
        Case "$B$10": suchSpalte = LagerMakros.Spalte_Finden(wsA, "MWST"):        titel = "MwSt"
        Case "$B$13": suchSpalte = LagerMakros.Spalte_Finden(wsA, "Attribut"):    titel = "Attribut"
    End Select
    If suchSpalte = 0 Then Exit Sub

    ' Eindeutige Treffer sammeln
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.count, 2).End(xlUp).Row
    Dim count As Integer: count = 0
    Dim i As Long

    For i = 3 To lastA
        If count >= 18 Then Exit For
        Dim cellVal As String: cellVal = Trim(CStr(wsA.Cells(i, suchSpalte).Value))
        If LCase(cellVal) Like "*" & suchbegriff & "*" And cellVal <> "" Then
            If Not dict.Exists(cellVal) Then
                dict.Add cellVal, cellVal
                wsNA.Cells(5 + count, 4).Value = cellVal
                wsNA.Cells(5 + count, 4).Interior.Color = RGB(220, 230, 241)
                count = count + 1
            End If
        End If
    Next i

    If count = 0 Then
        wsNA.Cells(5, 4).Value = "(Kein Treffer)"
        wsNA.Cells(5, 4).Font.Italic = True
        wsNA.Cells(5, 4).Font.Color = RGB(150, 150, 150)
    End If

    ' Merken welches Feld gerade aktiv ist
    wsNA.Cells(4, 4).Value = titel & " - Doppelklick zum Uebernehmen:"
    wsNA.Cells(4, 4).Font.Bold = True
    wsNA.Cells(4, 4).Font.Color = RGB(0, 70, 127)
End Sub


' ================================================================
'  Neuer Artikel - Vorschlag per Doppelklick uebernehmen
' ================================================================
Sub NeuerArtikel_VorschlagUebernehmen(Target As Range)
    Dim wsNA As Worksheet: Set wsNA = Target.Worksheet
    Dim val As String: val = Trim(CStr(Target.Value))
    If val = "" Or val = "(Kein Treffer)" Then Exit Sub

    Dim titel As String: titel = CStr(wsNA.Cells(4, 4).Value)

    Application.EnableEvents = False
    If InStr(titel, "Warengruppe") > 0 Then
        wsNA.Range("B8").Value = val
    ElseIf InStr(titel, "Lagerort") > 0 Then
        wsNA.Range("B9").Value = val
    ElseIf InStr(titel, "MwSt") > 0 Then
        wsNA.Range("B10").Value = val
    ElseIf InStr(titel, "Attribut") > 0 Then
        wsNA.Range("B13").Value = val
    End If
    Application.EnableEvents = True

    ' Vorschlaege leeren
    wsNA.Range("D5:D25").ClearContents
    wsNA.Range("D5:D25").Interior.ColorIndex = xlNone
    wsNA.Cells(4, 4).Value = "< Doppelklick auf Vorschlag zum Uebernehmen"
    wsNA.Cells(4, 4).Font.Bold = False
End Sub


' ================================================================
'  Neuer Artikel - Artikel speichern
' ================================================================
Sub NeuerArtikel_Speichern()
    Dim wsNA As Worksheet: Set wsNA = LagerMakros.GetSheet("Neuer Artikel")
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsNA Is Nothing Or wsA Is Nothing Then Exit Sub

    If Trim(CStr(wsNA.Range("B4").Value)) = "" Then
        MsgBox "Artikelnummer ist Pflicht!", vbExclamation
        wsNA.Range("B4").Select
        Exit Sub
    End If

    ' Spalten finden
    Dim cEAN  As Long: cEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim cArt  As Long: cArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim cVK   As Long: cVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim cEK   As Long: cEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim cMwst As Long: cMwst = LagerMakros.Spalte_Finden(wsA, "MWST")
    Dim cAnz  As Long: cAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim cEinh As Long: cEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim cLag  As Long: cLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim cWG   As Long: cWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim cNr   As Long: cNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim cAttr As Long: cAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")
    Dim cTA   As Long: cTA = LagerMakros.Spalte_Finden(wsA, "TextA")
    Dim cTB   As Long: cTB = LagerMakros.Spalte_Finden(wsA, "TextB")

    ' Neue Zeile
    Dim nextRow As Long: nextRow = wsA.Cells(wsA.Rows.count, cArt).End(xlUp).Row + 1

    ' Format von Zeile 3 kopieren (Farben/Format)
    On Error Resume Next
    wsA.Rows(3).Copy
    wsA.Rows(nextRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0

    ' Werte eintragen
    If cNr > 0 Then wsA.Cells(nextRow, cNr).Value = CStr(wsNA.Range("B4").Value)
    If cArt > 0 Then wsA.Cells(nextRow, cArt).Value = wsNA.Range("B5").Value
    If cEK > 0 Then wsA.Cells(nextRow, cEK).Value = val(wsNA.Range("B6").Value)
    If cVK > 0 Then wsA.Cells(nextRow, cVK).Value = val(wsNA.Range("B7").Value)
    If cWG > 0 Then wsA.Cells(nextRow, cWG).Value = wsNA.Range("B8").Value
    If cLag > 0 Then wsA.Cells(nextRow, cLag).Value = wsNA.Range("B9").Value
    If cMwst > 0 Then wsA.Cells(nextRow, cMwst).Value = val(wsNA.Range("B10").Value)
    If cEinh > 0 Then wsA.Cells(nextRow, cEinh).Value = wsNA.Range("B11").Value
    If cEAN > 0 Then
        wsA.Cells(nextRow, cEAN).NumberFormat = "@"
        wsA.Cells(nextRow, cEAN).Value = CStr(wsNA.Range("B12").Value)
    End If
    If cAttr > 0 Then wsA.Cells(nextRow, cAttr).Value = wsNA.Range("B13").Value
    If cTA > 0 Then wsA.Cells(nextRow, cTA).Value = wsNA.Range("B14").Value
    If cTB > 0 Then wsA.Cells(nextRow, cTB).Value = wsNA.Range("B15").Value
    If cAnz > 0 Then wsA.Cells(ne
