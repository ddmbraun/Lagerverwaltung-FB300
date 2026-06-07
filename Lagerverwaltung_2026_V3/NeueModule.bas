Attribute VB_Name = "NeueModule"
Option Explicit

' ================================================================
'  NEUE MODULE fuer 2026_Lagerverwaltung_V2.xlsm
'  Erstellt: 2026-06-06
'  Dieses Modul ergaenzt die fehlenden Funktionen:
'    - SchnellDetail (Popup beim Doppelklick in Schnellansicht)
'    - ArtikelDetail (Vollansicht mit Bearbeitungsfunktion)
'    - BewPopup     (Bewegungshistorie)
'    - InvDaten     (Inventurdaten-Speicher)
'    - InvEingabe   (Inventureingabe-Formular)
' ================================================================

Public g_AufruferSheet As String   ' Merkt wo ArtikelDetail aufgerufen wurde

' ================================================================
'  KONSTANTEN
' ================================================================
Const AD_ZEILE_REF  As Long = 17   ' Zeile in ArtikelDetail: speichert Artikelzeile
Const IE_ZEILE_REF  As Long = 14   ' Zeile in InvEingabe: speichert Artikelzeile
Const ID_DATEN_START As Long = 4   ' InvDaten: Datenzeilen ab Zeile 4

' ================================================================
'  ALLES EINRICHTEN - dieses Makro einmal ausfuehren!
'  Erstellt alle 5 fehlenden Sheets automatisch.
' ================================================================
Sub AllesEinrichten()
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim setupLog As String
    On Error GoTo SetupFehler

    setupLog = "SchnellDetail"
    Setup_SchnellDetail

    setupLog = "ArtikelDetail"
    Setup_ArtikelDetail

    setupLog = "BewPopup"
    Setup_BewPopup

    setupLog = "InvDaten"
    Setup_InvDaten

    setupLog = "InvEingabe"
    Setup_InvEingabe

    ' Doppelklick-Handler in Schnellansicht automatisch eintragen (optional)
    setupLog = "Schnellansicht-DoubleClick"
    Setup_Schnellansicht_DoubleClick

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Einrichtung abgeschlossen!" & Chr(10) & Chr(10) & _
           "Folgende Sheets wurden erstellt:" & Chr(10) & _
           "  - SchnellDetail" & Chr(10) & _
           "  - ArtikelDetail" & Chr(10) & _
           "  - BewPopup" & Chr(10) & _
           "  - InvDaten" & Chr(10) & _
           "  - InvEingabe" & Chr(10) & Chr(10) & _
           "Jetzt testen: Schnellansicht oeffnen und" & Chr(10) & _
           "einen Artikel doppelklicken.", _
           vbInformation, "Setup erfolgreich"
    Exit Sub

SetupFehler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Fehler bei: " & setupLog & Chr(10) & Chr(10) & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Setup-Fehler"
End Sub


' ================================================================
'  SETUP - Doppelklick-Handler in Schnellansicht eintragen
'  (erfordert: Extras > Makro > Sicherheit >
'   "Zugriff auf VBA-Projektobjektmodell vertrauen" = AN)
' ================================================================
Sub Setup_Schnellansicht_DoubleClick()
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub

    On Error GoTo OhneVBProject
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsS.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule

    ' Pruefen ob Artikelzeilen-Doppelklick schon vorhanden
    Dim i As Long
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "SchnellDetail_Laden") > 0 Then
            ' Schon vorhanden - nichts tun
            Exit Sub
        End If
    Next i

    ' Bestehende BeforeDoubleClick-Procedure finden und erweitern
    Dim gefunden As Long: gefunden = 0
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "BeforeDoubleClick") > 0 Then
            gefunden = i
            Exit For
        End If
    Next i

    If gefunden > 0 Then
        ' Zeile nach "Private Sub Worksheet_BeforeDoubleClick..." einfuegen
        Dim einfuegen As String
        einfuegen = "    If Target.Row >= 4 And Target.Column >= 2 And Target.Column <= 12 Then" & Chr(10) & _
                    "        Cancel = True" & Chr(10) & _
                    "        Dim sEAN As String: sEAN = CStr(Me.Cells(Target.Row, 5).Value)" & Chr(10) & _
                    "        Dim sArt As String: sArt = CStr(Me.Cells(Target.Row, 4).Value)" & Chr(10) & _
                    "        If sArt <> """" Then NeueModule.SchnellDetail_Laden sEAN, sArt" & Chr(10) & _
                    "    End If"
        cm.InsertLines gefunden + 1, einfuegen
    End If
    Exit Sub

OhneVBProject:
    ' Kein VBProject-Zugriff - Benutzer manuell anleiten (Schritt in Anleitung)
    Application.StatusBar = "INFO: Doppelklick-Handler muss manuell in Tabelle17 eingetragen werden."
End Sub


' ================================================================
'  SETUP - SchnellDetail Sheet erstellen
' ================================================================
Sub Setup_SchnellDetail()
    Dim wsSD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "SchnellDetail" Then Set wsSD = ws: Exit For
    Next ws
    If wsSD Is Nothing Then
        Set wsSD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsSD.Name = "SchnellDetail"
    End If

    wsSD.Cells.Clear
    wsSD.Cells.Interior.ColorIndex = xlNone

    ' Zeile 1: Titel
    On Error Resume Next: wsSD.Range("A1:E1").Merge: On Error GoTo 0
    wsSD.Cells(1, 1).Value = "Artikel - Schnellansicht"
    wsSD.Cells(1, 1).Interior.Color = RGB(31, 73, 125)
    wsSD.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsSD.Cells(1, 1).Font.Bold = True
    wsSD.Cells(1, 1).Font.Size = 14
    wsSD.Cells(1, 1).HorizontalAlignment = xlCenter
    wsSD.Rows(1).RowHeight = 30

    ' Zeile 2: Buttons (als Shapes)
    wsSD.Rows(2).RowHeight = 28

    ' Felder Zeile 3-12
    Dim felder As Variant
    felder = Array("Artikel:", "Art.-Nr.:", "EAN:", _
                   "VK-Preis " & ChrW(8364) & ":", _
                   "EK-Preis " & ChrW(8364) & ":", _
                   "Bestand:", "Einheit:", "Lagerort:", "Warengruppe:", "Attribut:")
    Dim r As Long
    For r = 0 To 9
        wsSD.Cells(r + 3, 2).Value = felder(r)
        wsSD.Cells(r + 3, 2).Font.Bold = True
        wsSD.Cells(r + 3, 2).Font.Size = 11
        wsSD.Cells(r + 3, 3).Interior.Color = RGB(255, 255, 200)
        wsSD.Cells(r + 3, 3).Font.Size = 12
        On Error Resume Next
        wsSD.Range(wsSD.Cells(r + 3, 3), wsSD.Cells(r + 3, 5)).Merge
        On Error GoTo 0
        wsSD.Rows(r + 3).RowHeight = 26
    Next r

    ' Spaltenbreiten
    wsSD.Columns(1).ColumnWidth = 2
    wsSD.Columns(2).ColumnWidth = 16
    wsSD.Columns(3).ColumnWidth = 20
    wsSD.Columns(4).ColumnWidth = 12
    wsSD.Columns(5).ColumnWidth = 8

    ' EAN-Feld als Text formatieren (verhindert Exponentialanzeige)
    wsSD.Cells(5, 3).NumberFormat = "@"

    ' Alte Shapes entfernen
    Dim shp As Shape
    For Each shp In wsSD.Shapes
        shp.Delete
    Next shp

    ' Button 1: EK EINBL. (blau)
    Dim rEK As Range: Set rEK = wsSD.Range("B2:C2")
    Dim oEK As Shape
    Set oEK = wsSD.Shapes.AddShape(msoShapeRectangle, _
        rEK.Left, rEK.Top, rEK.Width, rEK.Height)
    With oEK
        .Name = "btnEKToggle"
        .Fill.ForeColor.RGB = RGB(91, 155, 213)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "EK EINBL."
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "NeueModule.SchnellDetail_EK_Toggle"
    End With

    ' Button 2: SCHLIESSEN (rot)
    Dim rBtn As Range: Set rBtn = wsSD.Range("D2:E2")
    Dim oBtn As Shape
    Set oBtn = wsSD.Shapes.AddShape(msoShapeRectangle, _
        rBtn.Left, rBtn.Top, rBtn.Width, rBtn.Height)
    With oBtn
        .Name = "btnSchnellDetailSchliessen"
        .Fill.ForeColor.RGB = RGB(192, 0, 0)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "SCHLIESSEN"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "NeueModule.SchnellDetail_Schliessen"
    End With

    ' EK-Zeile (Zeile 7) standardmaessig ausblenden
    wsSD.Rows(7).Hidden = True

    wsSD.Visible = xlSheetHidden
End Sub


' ================================================================
'  SchnellDetail - Artikel laden und anzeigen
' ================================================================
Sub SchnellDetail_Laden(sEAN As String, sArtName As String)
    Dim wsSD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "SchnellDetail" Then Set wsSD = ws: Exit For
    Next ws
    If wsSD Is Nothing Then
        MsgBox "SchnellDetail-Sheet fehlt. Bitte 'AllesEinrichten' ausfuehren.", vbExclamation
        Exit Sub
    End If

    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub

    On Error GoTo SchnellDetailFehler

    Dim colEAN  As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colArt  As Long: colArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr   As Long: colNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim colVK   As Long: colVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK   As Long: colEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim colAnz  As Long: colAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim colEinh As Long: colEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim colLag  As Long: colLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim colWG   As Long: colWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colAttr As Long: colAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")

    If colArt = 0 Then GoTo SchnellDetailOeffnen

    Dim lastRow As Long: lastRow = wsA.Cells(wsA.Rows.count, colArt).End(xlUp).Row
    Dim gefunden As Long: gefunden = 0
    Dim i As Long

    ' Zuerst per EAN suchen, dann per Name
    If sEAN <> "" And colEAN > 0 Then
        For i = 3 To lastRow
            If CStr(wsA.Cells(i, colEAN).Value) = sEAN Then
                gefunden = i: Exit For
            End If
        Next i
    End If
    If gefunden = 0 And sArtName <> "" Then
        For i = 3 To lastRow
            If CStr(wsA.Cells(i, colArt).Value) = sArtName Then
                gefunden = i: Exit For
            End If
        Next i
    End If

    If gefunden > 0 Then
        Dim vNr   As String: If colNr > 0 Then vNr = CStr(wsA.Cells(gefunden, colNr).Value)
        Dim vVK   As Double: If colVK > 0 Then vVK = val(wsA.Cells(gefunden, colVK).Value)
        Dim vEK   As Double: If colEK > 0 Then vEK = val(wsA.Cells(gefunden, colEK).Value)
        Dim vAnz  As Double: If colAnz > 0 Then vAnz = val(wsA.Cells(gefunden, colAnz).Value)
        Dim vEinh As String: If colEinh > 0 Then vEinh = CStr(wsA.Cells(gefunden, colEinh).Value)
        Dim vLag  As String: If colLag > 0 Then vLag = CStr(wsA.Cells(gefunden, colLag).Value)
        Dim vWG   As String: If colWG > 0 Then vWG = CStr(wsA.Cells(gefunden, colWG).Value)
        Dim vAttr As String: If colAttr > 0 Then vAttr = CStr(wsA.Cells(gefunden, colAttr).Value)

        ' EAN als Text formatiert speichern
        Dim vEANText As String
        If colEAN > 0 Then
            Dim rawEAN As Variant: rawEAN = wsA.Cells(gefunden, colEAN).Value
            If IsNumeric(rawEAN) Then
                vEANText = Format(CDbl(rawEAN), "0000000000000")
            Else
                vEANText = CStr(rawEAN)
            End If
        End If

        ' Felder befuellen
        wsSD.Cells(3, 3).Value = sArtName
        wsSD.Cells(4, 3).Value = vNr
        wsSD.Cells(5, 3).NumberFormat = "@"
        wsSD.Cells(5, 3).Value = vEANText
        wsSD.Cells(6, 3).Value = vVK
        wsSD.Cells(6, 3).NumberFormat = "#,##0.00"
        wsSD.Cells(7, 3).Value = vEK
        wsSD.Cells(7, 3).NumberFormat = "#,##0.00"
        wsSD.Cells(8, 3).Value = vAnz
        wsSD.Cells(8, 3).NumberFormat = "0"
        wsSD.Cells(9, 3).Value = vEinh
        wsSD.Cells(10, 3).Value = vLag
        wsSD.Cells(11, 3).Value = vWG
        wsSD.Cells(12, 3).Value = vAttr
    End If

SchnellDetailOeffnen:
    wsSD.Visible = xlSheetVisible
    wsSD.Activate
    Exit Sub

SchnellDetailFehler:
    MsgBox "Fehler " & Err.Number & " in SchnellDetail_Laden: " & Err.Description, vbCritical
End Sub


' ================================================================
'  SchnellDetail - Schliessen -> zurueck zur Schnellansicht
' ================================================================
Sub SchnellDetail_Schliessen()
    Dim wsSD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "SchnellDetail" Then Set wsSD = ws: Exit For
    Next ws
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If Not wsS Is Nothing Then
        wsS.Visible = xlSheetVisible
        wsS.Activate
    End If
    If Not wsSD Is Nothing Then wsSD.Visible = xlSheetHidden
End Sub


' ================================================================
'  SchnellDetail - EK-Preis ein-/ausblenden
' ================================================================
Sub SchnellDetail_EK_Toggle()
    Dim wsSD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "SchnellDetail" Then Set wsSD = ws: Exit For
    Next ws
    If wsSD Is Nothing Then Exit Sub

    Dim istVersteckt As Boolean: istVersteckt = wsSD.Rows(7).Hidden
    wsSD.Rows(7).Hidden = Not istVersteckt

    Dim shp As Shape
    For Each shp In wsSD.Shapes
        If shp.Name = "btnEKToggle" Then
            If wsSD.Rows(7).Hidden Then
                shp.TextFrame.Characters.Text = "EK EINBL."
            Else
                shp.TextFrame.Characters.Text = "EK AUSBL."
            End If
            Exit For
        End If
    Next shp
End Sub


' ================================================================
'  SETUP - ArtikelDetail Sheet erstellen
' ================================================================
Sub Setup_ArtikelDetail()
    Dim wsAD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "ArtikelDetail" Then Set wsAD = ws: Exit For
    Next ws
    If wsAD Is Nothing Then
        Set wsAD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsAD.Name = "ArtikelDetail"
    End If

    Application.ScreenUpdating = False
    wsAD.Cells.Clear
    wsAD.Cells.Interior.ColorIndex = xlNone

    Dim blau As Long:     blau = RGB(32, 55, 100)
    Dim gruen As Long:    gruen = RGB(55, 110, 50)
    Dim orange As Long:   orange = RGB(180, 90, 0)
    Dim gelb As Long:     gelb = RGB(255, 255, 0)
    Dim hellgrau As Long: hellgrau = RGB(242, 242, 242)

    ' Zeile 1: Titel
    On Error Resume Next: wsAD.Range("A1:G1").Merge: On Error GoTo 0
    wsAD.Cells(1, 1).Value = "Artikel - Details"
    wsAD.Cells(1, 1).Interior.Color = blau
    wsAD.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsAD.Cells(1, 1).Font.Size = 14
    wsAD.Cells(1, 1).Font.Bold = True
    wsAD.Cells(1, 1).HorizontalAlignment = xlCenter
    wsAD.Rows(1).RowHeight = 30

    ' Felder: Zeile, LinksBeschrift, RechsBeschrift
    Dim felder As Variant
    felder = Array( _
        Array(3, "Artikel:", ""), _
        Array(4, "Art.-Nr.:", "EAN:"), _
        Array(5, "VK-Preis:", "VK2 (opt.):"), _
        Array(6, "EK-Preis:", ""), _
        Array(7, "MwSt %:", "Lieferant:"), _
        Array(8, "Bestand:", "Einheit:"), _
        Array(9, "Lagerort:", "Warengruppe:"), _
        Array(10, "Attribut:", ""), _
        Array(11, "TextA:", ""), _
        Array(12, "TextB:", ""), _
        Array(13, "Rohgewinn:", "") _
    )

    Dim k As Integer
    For k = 0 To UBound(felder)
        Dim r As Long: r = felder(k)(0)
        wsAD.Cells(r, 2).Value = felder(k)(1)
        wsAD.Cells(r, 2).Font.Bold = True
        wsAD.Cells(r, 2).Font.Size = 11
        wsAD.Cells(r, 3).Interior.Color = gelb
        wsAD.Cells(r, 3).Font.Size = 11
        On Error Resume Next: wsAD.Range(wsAD.Cells(r, 3), wsAD.Cells(r, 4)).Merge: On Error GoTo 0
        If felder(k)(2) <> "" Then
            wsAD.Cells(r, 5).Value = felder(k)(2)
            wsAD.Cells(r, 5).Font.Bold = True
            wsAD.Cells(r, 5).Font.Size = 11
            wsAD.Cells(r, 6).Interior.Color = gelb
            wsAD.Cells(r, 6).Font.Size = 11
            On Error Resume Next: wsAD.Range(wsAD.Cells(r, 6), wsAD.Cells(r, 7)).Merge: On Error GoTo 0
        End If
        wsAD.Rows(r).RowHeight = 22
    Next k

    ' Zeile 3: Artikelname breiter
    On Error Resume Next: wsAD.Range("C3:G3").Merge: On Error GoTo 0
    wsAD.Rows(3).RowHeight = 26

    ' Zeile 13: Rohgewinn - nur anzeigen, nicht bearbeiten
    wsAD.Cells(13, 3).Interior.Color = hellgrau
    wsAD.Cells(13, 3).Font.Bold = True

    ' Zeile 14: Leerzeile
    wsAD.Rows(14).RowHeight = 8

    ' Zeile 15: Buttons
    With wsAD.Cells(15, 2)
        .Value = "SCHLIESSEN"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(89, 89, 89): .HorizontalAlignment = xlCenter
    End With
    On Error Resume Next: wsAD.Range("B15:C15").Merge: On Error GoTo 0
    wsAD.Rows(15).RowHeight = 28

    With wsAD.Cells(15, 4)
        .Value = "BEWEGUNGEN"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = orange: .HorizontalAlignment = xlCenter
    End With
    On Error Resume Next: wsAD.Range("D15:E15").Merge: On Error GoTo 0

    With wsAD.Cells(15, 6)
        .Value = "SPEICHERN"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = gruen: .HorizontalAlignment = xlCenter
    End With
    On Error Resume Next: wsAD.Range("F15:G15").Merge: On Error GoTo 0

    ' Zeile 16: Leerzeile
    wsAD.Rows(16).RowHeight = 6

    ' Zeile 17: versteckt - speichert Artikelzeile-Nummer
    wsAD.Rows(17).Hidden = True
    wsAD.Cells(AD_ZEILE_REF, 1).Value = 0

    ' Spaltenbreiten
    wsAD.Columns(1).ColumnWidth = 2
    wsAD.Columns(2).ColumnWidth = 16
    wsAD.Columns(3).ColumnWidth = 20
    wsAD.Columns(4).ColumnWidth = 16
    wsAD.Columns(5).ColumnWidth = 16
    wsAD.Columns(6).ColumnWidth = 20
    wsAD.Columns(7).ColumnWidth = 10

    wsAD.Visible = xlSheetHidden
    Application.ScreenUpdating = True
End Sub


' ================================================================
'  ArtikelDetail - Artikel laden (aus Artikel-Sheet Zeile)
' ================================================================
Sub ArtikelDetail_Laden(zeile As Long)
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    If wsA Is Nothing Or wsAD Is Nothing Then
        MsgBox "ArtikelDetail-Blatt fehlt. Bitte 'AllesEinrichten' ausfuehren.", vbExclamation
        Exit Sub
    End If

    ' Aufrufer merken
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If Not wsS Is Nothing And ActiveSheet.Name = wsS.Name Then
        g_AufruferSheet = "Schnellansicht"
    Else
        g_AufruferSheet = "Artikel"
    End If

    On Error GoTo ArtikelDetailFehler

    Dim colArt  As Long: colArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr   As Long: colNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim colEAN  As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colVK   As Long: colVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK   As Long: colEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim colMwst As Long: colMwst = LagerMakros.Spalte_Finden(wsA, "MWST")
    Dim colLief As Long: colLief = LagerMakros.Spalte_Finden(wsA, "LIEFERANT")
    Dim colAnz  As Long: colAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim colEinh As Long: colEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim colLag  As Long: colLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim colWG   As Long: colWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colAttr As Long: colAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")
    Dim colTA   As Long: colTA = LagerMakros.Spalte_Finden(wsA, "TextA")
    Dim colTB   As Long: colTB = LagerMakros.Spalte_Finden(wsA, "TextB")
    Dim colVK2  As Long: colVK2 = LagerMakros.Spalte_Finden(wsA, "VK2")

    Dim vArt  As String: If colArt > 0 Then vArt = CStr(wsA.Cells(zeile, colArt).Value)
    Dim vNr   As String: If colNr > 0 Then vNr = CStr(wsA.Cells(zeile, colNr).Value)
    Dim vLief As String: If colLief > 0 Then vLief = CStr(wsA.Cells(zeile, colLief).Value)
    Dim vEinh As String: If colEinh > 0 Then vEinh = CStr(wsA.Cells(zeile, colEinh).Value) Else vEinh = "Stk"
    Dim vLag  As String: If colLag > 0 Then vLag = CStr(wsA.Cells(zeile, colLag).Value)
    Dim vWG   As String: If colWG > 0 Then vWG = CStr(wsA.Cells(zeile, colWG).Value)
    Dim vAttr As String: If colAttr > 0 Then vAttr = CStr(wsA.Cells(zeile, colAttr).Value)
    Dim vTA   As String: If colTA > 0 Then vTA = CStr(wsA.Cells(zeile, colTA).Value)
    Dim vTB   As String: If colTB > 0 Then vTB = CStr(wsA.Cells(zeile, colTB).Value)
    Dim vVK   As Double: If colVK > 0 Then vVK = val(wsA.Cells(zeile, colVK).Value)
    Dim vEK   As Double: If colEK > 0 Then vEK = val(wsA.Cells(zeile, colEK).Value)
    Dim vMwst As Double: If colMwst > 0 Then vMwst = val(wsA.Cells(zeile, colMwst).Value) Else vMwst = 19
    Dim vAnz  As Double: If colAnz > 0 Then vAnz = val(wsA.Cells(zeile, colAnz).Value)
    Dim vVK2  As Double: If colVK2 > 0 Then vVK2 = val(wsA.Cells(zeile, colVK2).Value)

    ' EAN als Text
    Dim vEAN As String
    If colEAN > 0 Then
        Dim rawEAN As Variant: rawEAN = wsA.Cells(zeile, colEAN).Value
        If IsNumeric(rawEAN) Then
            vEAN = Format(CDbl(rawEAN), "0000000000000")
        Else
            vEAN = CStr(rawEAN)
        End If
    End If

    ' Rohgewinn berechnen
    Dim rohgewinn As Double: rohgewinn = vVK - vEK
    Dim marge As Double: If vVK > 0 Then marge = rohgewinn / vVK * 100
    Dim vRG As String: vRG = Format(rohgewinn, "0.00") & " EUR  (" & Format(marge, "0.0") & " %)"

    ' In ArtikelDetail eintragen
    wsAD.Cells(3, 3).Value = vArt
    wsAD.Cells(4, 3).Value = vNr
    wsAD.Cells(4, 6).NumberFormat = "@"
    wsAD.Cells(4, 6).Value = vEAN
    wsAD.Cells(5, 3).Value = vVK
    wsAD.Cells(5, 3).NumberFormat = "0.00"
    wsAD.Cells(5, 6).Value = IIf(vVK2 > 0, vVK2, "")
    wsAD.Cells(6, 3).Value = vEK
    wsAD.Cells(6, 3).NumberFormat = "0.00"
    wsAD.Cells(7, 3).Value = vMwst
    wsAD.Cells(7, 6).Value = vLief
    wsAD.Cells(8, 3).Value = vAnz
    wsAD.Cells(8, 6).Value = vEinh
    wsAD.Cells(9, 3).Value = vLag
    wsAD.Cells(9, 6).Value = vWG
    wsAD.Cells(10, 3).Value = vAttr
    wsAD.Cells(11, 3).Value = vTA
    wsAD.Cells(12, 3).Value = vTB
    wsAD.Cells(13, 3).Value = vRG
    wsAD.Cells(AD_ZEILE_REF, 1).Value = zeile

    wsAD.Visible = xlSheetVisible
    wsAD.Activate
    On Error Resume Next: wsAD.Cells(3, 3).Select: On Error GoTo 0
    Exit Sub

ArtikelDetailFehler:
    MsgBox "Fehler " & Err.Number & " in ArtikelDetail_Laden: " & Err.Description, vbCritical
    g_AufruferSheet = ""
End Sub


' ================================================================
'  ArtikelDetail - Speichern -> Daten zurueck ins Artikel-Sheet
' ================================================================
Sub ArtikelDetail_Speichern()
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    If wsA Is Nothing Or wsAD Is Nothing Then Exit Sub

    Dim zeile As Long: zeile = val(wsAD.Cells(AD_ZEILE_REF, 1).Value)
    If zeile < 3 Then MsgBox "Kein Artikel geladen.", vbExclamation: Exit Sub

    Dim colArt  As Long: colArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr   As Long: colNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim colEAN  As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colVK   As Long: colVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK   As Long: colEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim colMwst As Long: colMwst = LagerMakros.Spalte_Finden(wsA, "MWST")
    Dim colLief As Long: colLief = LagerMakros.Spalte_Finden(wsA, "LIEFERANT")
    Dim colAnz  As Long: colAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim colEinh As Long: colEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim colLag  As Long: colLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim colWG   As Long: colWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colAttr As Long: colAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")
    Dim colTA   As Long: colTA = LagerMakros.Spalte_Finden(wsA, "TextA")
    Dim colTB   As Long: colTB = LagerMakros.Spalte_Finden(wsA, "TextB")
    Dim colVK2  As Long: colVK2 = LagerMakros.Spalte_Finden(wsA, "VK2")

    If colArt > 0 Then wsA.Cells(zeile, colArt).Value = wsAD.Cells(3, 3).Value
    If colNr > 0 Then wsA.Cells(zeile, colNr).Value = CStr(wsAD.Cells(4, 3).Value)
    If colEAN > 0 Then
        wsA.Cells(zeile, colEAN).NumberFormat = "@"
        wsA.Cells(zeile, colEAN).Value = CStr(wsAD.Cells(4, 6).Value)
    End If
    If colVK > 0 Then wsA.Cells(zeile, colVK).Value = val(wsAD.Cells(5, 3).Value)
    If colVK2 > 0 Then wsA.Cells(zeile, colVK2).Value = val(wsAD.Cells(5, 6).Value)
    If colEK > 0 Then wsA.Cells(zeile, colEK).Value = val(wsAD.Cells(6, 3).Value)
    If colMwst > 0 Then wsA.Cells(zeile, colMwst).Value = val(wsAD.Cells(7, 3).Value)
    If colLief > 0 Then wsA.Cells(zeile, colLief).Value = wsAD.Cells(7, 6).Value
    If colAnz > 0 Then wsA.Cells(zeile, colAnz).Value = val(wsAD.Cells(8, 3).Value)
    If colEinh > 0 Then wsA.Cells(zeile, colEinh).Value = wsAD.Cells(8, 6).Value
    If colLag > 0 Then wsA.Cells(zeile, colLag).Value = wsAD.Cells(9, 3).Value
    If colWG > 0 Then wsA.Cells(zeile, colWG).Value = wsAD.Cells(9, 6).Value
    If colAttr > 0 Then wsA.Cells(zeile, colAttr).Value = wsAD.Cells(10, 3).Value
    If colTA > 0 Then wsA.Cells(zeile, colTA).Value = wsAD.Cells(11, 3).Value
    If colTB > 0 Then wsA.Cells(zeile, colTB).Value = wsAD.Cells(12, 3).Value

    MsgBox "Artikel gespeichert.", vbInformation
    ArtikelDetail_Schliessen
End Sub


' ================================================================
'  ArtikelDetail - Schliessen -> zurueck zum Aufrufer
' ================================================================
Sub ArtikelDetail_Schliessen()
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    Dim zeile As Long
    If Not wsAD Is Nothing Then
        zeile = val(wsAD.Cells(AD_ZEILE_REF, 1).Value)
        wsAD.Visible = xlSheetHidden
    End If
    Application.StatusBar = False
    If g_AufruferSheet = "Schnellansicht" Then
        Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
        If Not wsS Is Nothing Then
            wsS.Visible = xlSheetVisible
            wsS.Activate
        End If
    Else
        If Not wsA Is Nothing Then
            wsA.Activate
            If zeile >= 3 Then Application.GoTo wsA.Cells(zeile, 1), True
        End If
    End If
    g_AufruferSheet = ""
End Sub


' ================================================================
'  ArtikelDetail - Bewegungshistorie oeffnen
' ================================================================
Sub ArtikelDetail_Bewegungen()
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    If wsAD Is Nothing Then Exit Sub
    Dim ean     As String: ean = CStr(wsAD.Cells(4, 6).Value)
    Dim artName As String: artName = CStr(wsAD.Cells(3, 3).Value)
    BewPopup_Laden ean, artName
End Sub


' ================================================================
'  SETUP - BewPopup Sheet erstellen
' ================================================================
Sub Setup_BewPopup()
    Dim wsBP As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "BewPopup" Then Set wsBP = ws: Exit For
    Next ws
    If wsBP Is Nothing Then
        Set wsBP = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsBP.Name = "BewPopup"
    End If

    Application.ScreenUpdating = False
    wsBP.Cells.Clear
    wsBP.Cells.Interior.ColorIndex = xlNone

    ' Zeile 1: Titel
    On Error Resume Next: wsBP.Range("A1:F1").Merge: On Error GoTo 0
    wsBP.Cells(1, 1).Value = "Bewegungshistorie"
    wsBP.Cells(1, 1).Interior.Color = RGB(31, 56, 100)
    wsBP.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsBP.Cells(1, 1).Font.Bold = True
    wsBP.Cells(1, 1).Font.Size = 13
    wsBP.Cells(1, 1).HorizontalAlignment = xlCenter
    wsBP.Rows(1).RowHeight = 28

    ' Zeile 2: Artikelname + Schliessen-Button
    On Error Resume Next: wsBP.Range("B2:E2").Merge: On Error GoTo 0
    wsBP.Cells(2, 2).Font.Bold = True
    wsBP.Cells(2, 2).Font.Size = 11
    With wsBP.Cells(2, 6)
        .Value = "SCHLIESSEN"
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlCenter
    End With
    wsBP.Rows(2).RowHeight = 24

    ' Zeile 3: Spaltenkoepfe
    Dim hdrs As Variant
    hdrs = Array("Datum/Zeit", "EAN", "Menge", "Typ", "Lagerort", "Benutzer")
    Dim j As Integer
    For j = 0 To 5
        wsBP.Cells(3, j + 1).Value = hdrs(j)
        wsBP.Cells(3, j + 1).Interior.Color = RGB(46, 80, 144)
        wsBP.Cells(3, j + 1).Font.Color = RGB(255, 255, 255)
        wsBP.Cells(3, j + 1).Font.Bold = True
    Next j
    wsBP.Rows(3).RowHeight = 20

    ' Spaltenbreiten
    wsBP.Columns(1).ColumnWidth = 16
    wsBP.Columns(2).ColumnWidth = 14
    wsBP.Columns(3).ColumnWidth = 8
    wsBP.Columns(4).ColumnWidth = 10
    wsBP.Columns(5).ColumnWidth = 14
    wsBP.Columns(6).ColumnWidth = 12

    wsBP.Visible = xlSheetHidden
    Application.ScreenUpdating = True
End Sub


' ================================================================
'  BewPopup - Bewegungen laden und anzeigen
' ================================================================
Sub BewPopup_Laden(ean As String, artName As String)
    Dim wsBP As Worksheet: Set wsBP = LagerMakros.GetSheet("BewPopup")
    Dim wsZ  As Worksheet: Set wsZ = LagerMakros.GetSheet("Abg")
    If wsBP Is Nothing Then
        MsgBox "BewPopup-Sheet fehlt. Bitte 'AllesEinrichten' ausfuehren.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Alte Daten loeschen
    Dim lastBP As Long: lastBP = wsBP.Cells(wsBP.Rows.count, 1).End(xlUp).Row
    If lastBP >= 4 Then wsBP.Range("A4:F" & lastBP).Clear

    ' Artikelname in Zeile 2
    wsBP.Cells(2, 2).Value = Left(artName, 40) & "  [EAN: " & ean & "]"

    ' Bewegungen suchen und anzeigen
    If Not wsZ Is Nothing Then
        Dim colEAN_Z As Long: colEAN_Z = LagerMakros.Spalte_Finden(wsZ, "EAN13")
        Dim lastZ As Long: lastZ = wsZ.Cells(wsZ.Rows.count, 1).End(xlUp).Row
        Dim sRow As Long: sRow = 4
        Dim i As Long
        For i = 2 To lastZ
            If colEAN_Z > 0 Then
                If CStr(wsZ.Cells(i, colEAN_Z).Value) = ean Then
                    wsBP.Cells(sRow, 1).Value = wsZ.Cells(i, 1).Value  ' Datum
                    wsBP.Cells(sRow, 1).NumberFormat = "DD.MM.YYYY HH:MM"
                    wsBP.Cells(sRow, 2).Value = ean
                    wsBP.Cells(sRow, 3).Value = wsZ.Cells(i, 5).Value  ' Menge
                    wsBP.Cells(sRow, 4).Value = wsZ.Cells(i, 6).Value  ' Typ
                    wsBP.Cells(sRow, 5).Value = wsZ.Cells(i, 7).Value  ' Lagerort
                    wsBP.Cells(sRow, 6).Value = wsZ.Cells(i, 8).Value  ' Benutzer
                    ' Farbe: Zugang=gruen, Abgang=rot
                    If CStr(wsZ.Cells(i, 6).Value) = "Zugang" Then
                        wsBP.Rows(sRow).Interior.Color = RGB(198, 239, 206)
                    Else
                        wsBP.Rows(sRow).Interior.Color = RGB(255, 199, 206)
                    End If
                    sRow = sRow + 1
                End If
            End If
        Next i
        If sRow = 4 Then
            wsBP.Cells(4, 1).Value = "(Keine Bewegungen gefunden)"
        End If
    End If

    ' ArtikelDetail ausblenden
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    If Not wsAD Is Nothing Then wsAD.Visible = xlSheetHidden

    wsBP.Visible = xlSheetVisible
    wsBP.Activate
    Application.ScreenUpdating = True
End Sub


' ================================================================
'  BewPopup - Schliessen -> zurueck zu ArtikelDetail
' ================================================================
Sub BewPopup_Schliessen()
    Dim wsBP As Worksheet: Set wsBP = LagerMakros.GetSheet("BewPopup")
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    If Not wsAD Is Nothing Then
        wsAD.Visible = xlSheetVisible
        wsAD.Activate
    Else
        Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
        If Not wsA Is Nothing Then wsA.Activate
    End If
    If Not wsBP Is Nothing Then wsBP.Visible = xlSheetHidden
End Sub


' ================================================================
'  SETUP - InvDaten Sheet erstellen (Inventur-Datenspeicher)
' ================================================================
Sub Setup_InvDaten()
    Dim wsID As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "InvDaten" Then Set wsID = ws: Exit For
    Next ws
    If wsID Is Nothing Then
        Set wsID = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsID.Name = "InvDaten"
    End If

    ' Header nur setzen wenn leer
    If wsID.Cells(1, 1).Value = "" Then
        wsID.Cells(1, 1).Value = "Inventur-Start:"
        wsID.Cells(1, 2).Value = Now()
        wsID.Cells(1, 2).NumberFormat = "DD.MM.YYYY HH:MM"
    End If

    If wsID.Cells(3, 1).Value = "" Then
        Dim hdrs As Variant
        hdrs = Array("Datum/Zeit", "EAN", "Art.-Nr.", "Artikel", "SOLL", "IST", "DIFFERENZ", "Lagerort")
        Dim j As Integer
        For j = 0 To 7
            wsID.Cells(3, j + 1).Value = hdrs(j)
            wsID.Cells(3, j + 1).Interior.Color = RGB(46, 80, 144)
            wsID.Cells(3, j + 1).Font.Color = RGB(255, 255, 255)
            wsID.Cells(3, j + 1).Font.Bold = True
        Next j
    End If
End Sub


' ================================================================
'  InvDaten - Eintrag speichern oder aktualisieren
' ================================================================
Sub InvDaten_Speichern(ean As String, artNr As String, artName As String, _
                       soll As Double, ist As Double, lagerort As String)
    Dim wsID As Worksheet: Set wsID = LagerMakros.GetSheet("InvDaten")
    If wsID Is Nothing Then
        MsgBox "InvDaten-Sheet fehlt.", vbExclamation
        Exit Sub
    End If

    Dim lastID As Long: lastID = wsID.Cells(wsID.Rows.count, 2).End(xlUp).Row
    Dim i As Long

    ' Pruefen ob EAN bereits vorhanden
    For i = ID_DATEN_START To lastID
        If CStr(wsID.Cells(i, 2).Value) = ean Then
            wsID.Cells(i, 1).Value = Now()
            wsID.Cells(i, 6).Value = ist
            wsID.Cells(i, 7).Value = ist - soll
            Exit Sub
        End If
    Next i

    ' Neuen Eintrag anfuegen
    Dim nRow As Long: nRow = IIf(lastID < ID_DATEN_START, ID_DATEN_START, lastID + 1)
    wsID.Cells(nRow, 1).Value = Now()
    wsID.Cells(nRow, 1).NumberFormat = "DD.MM.YYYY HH:MM:SS"
    wsID.Cells(nRow, 2).Value = ean
    wsID.Cells(nRow, 3).Value = artNr
    wsID.Cells(nRow, 4).Value = artName
    wsID.Cells(nRow, 5).Value = soll
    wsID.Cells(nRow, 6).Value = ist
    wsID.Cells(nRow, 7).Value = ist - soll
    wsID.Cells(nRow, 8).Value = lagerort
End Sub


' ================================================================
'  SETUP - InvEingabe Sheet erstellen
' ================================================================
Sub Setup_InvEingabe()
    Dim wsIE As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "InvEingabe" Then Set wsIE = ws: Exit For
    Next ws
    If wsIE Is Nothing Then
        Set wsIE = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsIE.Name = "InvEingabe"
    End If

    Application.ScreenUpdating = False
    wsIE.Cells.Clear
    wsIE.Cells.Interior.ColorIndex = xlNone

    Dim blau As Long:     blau = RGB(32, 55, 100)
    Dim gruen As Long:    gruen = RGB(55, 110, 50)
    Dim gelb As Long:     gelb = RGB(255, 255, 200)
    Dim hellgrau As Long: hellgrau = RGB(242, 242, 242)

    ' Zeile 1: Titel
    On Error Resume Next: wsIE.Range("A1:F1").Merge: On Error GoTo 0
    wsIE.Cells(1, 1).Value = "Inventur - Artikel pruefen"
    wsIE.Cells(1, 1).Interior.Color = blau
    wsIE.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsIE.Cells(1, 1).Font.Size = 14
    wsIE.Cells(1, 1).Font.Bold = True
    wsIE.Cells(1, 1).HorizontalAlignment = xlCenter
    wsIE.Rows(1).RowHeight = 30

    ' Felder Zeile 3-7
    Dim felder As Variant
    felder = Array( _
        Array(3, "Artikel:", ""), _
        Array(4, "Art.-Nr.:", "EAN:"), _
        Array(5, "VK-Preis:", "EK-Preis:"), _
        Array(6, "Lagerort:", "Warengruppe:"), _
        Array(7, "Attribut:", "") _
    )

    Dim k As Integer
    For k = 0 To UBound(felder)
        Dim r As Long: r = felder(k)(0)
        wsIE.Cells(r, 2).Value = felder(k)(1)
        wsIE.Cells(r, 2).Font.Bold = True
        wsIE.Cells(r, 2).Font.Size = 11
        wsIE.Cells(r, 3).Interior.Color = hellgrau
        wsIE.Cells(r, 3).Font.Size = 11
        On Error Resume Next: wsIE.Range(wsIE.Cells(r, 3), wsIE.Cells(r, 4)).Merge: On Error GoTo 0
        If felder(k)(2) <> "" Then
            wsIE.Cells(r, 5).Value = felder(k)(2)
            wsIE.Cells(r, 5).Font.Bold = True
            wsIE.Cells(r, 6).Interior.Color = hellgrau
        End If
        wsIE.Rows(r).RowHeight = 22
    Next k

    ' Zeile 3: Artikelname breiter
    On Error Resume Next: wsIE.Range("C3:F3").Merge: On Error GoTo 0

    ' Zeile 8: Trennzeile
    wsIE.Rows(8).RowHeight = 8

    ' Zeile 9: SOLL / GEZAEHLT
    wsIE.Cells(9, 2).Value = "SOLL:"
    wsIE.Cells(9, 2).Font.Bold = True: wsIE.Cells(9, 2).Font.Size = 12
    wsIE.Cells(9, 3).Interior.Color = hellgrau
    wsIE.Cells(9, 3).Font.Size = 12: wsIE.Cells(9, 3).Font.Bold = True
    On Error Resume Next: wsIE.Range("C9:D9").Merge: On Error GoTo 0
    wsIE.Cells(9, 5).Value = "GEZAEHLT:"
    wsIE.Cells(9, 5).Font.Bold = True: wsIE.Cells(9, 5).Font.Size = 12
    wsIE.Cells(9, 6).Interior.Color = RGB(255, 255, 0)
    wsIE.Cells(9, 6).Font.Size = 14: wsIE.Cells(9, 6).Font.Bold = True
    wsIE.Rows(9).RowHeight = 26

    ' Zeile 10: Differenz
    wsIE.Cells(10, 5).Value = "DIFFERENZ:"
    wsIE.Cells(10, 5).Font.Bold = True
    wsIE.Cells(10, 6).Font.Bold = True
    wsIE.Rows(10).RowHeight = 20

    ' Zeile 11: Buttons
    With wsIE.Cells(11, 2)
        .Value = "ABBRECHEN"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(89, 89, 89): .HorizontalAlignment = xlCenter
    End With
    On Error Resume Next: wsIE.Range("B11:C11").Merge: On Error GoTo 0
    With wsIE.Cells(11, 5)
        .Value = "UEBERNEHMEN"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = gruen: .HorizontalAlignment = xlCenter
    End With
    On Error Resume Next: wsIE.Range("E11:F11").Merge: On Error GoTo 0
    wsIE.Rows(11).RowHeight = 28

    ' Zeile 14: versteckt - speichert Artikelzeile
    wsIE.Rows(14).Hidden = True
    wsIE.Cells(IE_ZEILE_REF, 1).Value = 0

    ' Spaltenbreiten
    wsIE.Columns(1).ColumnWidth = 2
    wsIE.Columns(2).ColumnWidth = 14
    wsIE.Columns(3).ColumnWidth = 20
    wsIE.Columns(4).ColumnWidth = 10
    wsIE.Columns(5).ColumnWidth = 14
    wsIE.Columns(6).ColumnWidth = 20

    wsIE.Visible = xlSheetHidden
    Application.ScreenUpdating = True
End Sub


' ================================================================
'  InvEingabe - Artikel befuellen (aus InvSuche aufgerufen)
' ================================================================
Sub InvEingabe_Befuellen(zeile As Long)
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsIE As Worksheet: Set wsIE = LagerMakros.GetSheet("InvEingabe")
    If wsA Is Nothing Or wsIE Is Nothing Then
        MsgBox "InvEingabe-Blatt fehlt. Bitte 'AllesEinrichten' ausfuehren.", vbExclamation
        Exit Sub
    End If

    Dim colArt  As Long: colArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr   As Long: colNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim colEAN  As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colVK   As Long: colVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK   As Long: colEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim colAnz  As Long: colAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim colEinh As Long: colEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim colLag  As Long: colLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim colWG   As Long: colWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colAttr As Long: colAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")

    Dim vEAN As String
    If colEAN > 0 Then
        Dim rawEAN As Variant: rawEAN = wsA.Cells(zeile, colEAN).Value
        If IsNumeric(rawEAN) Then
            vEAN = Format(CDbl(rawEAN), "0000000000000")
        Else
            vEAN = CStr(rawEAN)
        End If
    End If

    Dim vAnz  As Double: If colAnz > 0 Then vAnz = val(wsA.Cells(zeile, colAnz).Value)
    Dim vEinh As String: If colEinh > 0 Then vEinh = CStr(wsA.Cells(zeile, colEinh).Value)

    wsIE.Cells(3, 3).Value = IIf(colArt > 0, wsA.Cells(zeile, colArt).Value, "")
    wsIE.Cells(4, 3).Value = IIf(colNr > 0, CStr(wsA.Cells(zeile, colNr).Value), "")
    wsIE.Cells(4, 6).Value = vEAN
    wsIE.Cells(5, 3).Value = IIf(colVK > 0, val(wsA.Cells(zeile, colVK).Value), "")
    wsIE.Cells(5, 6).Value = IIf(colEK > 0, val(wsA.Cells(zeile, colEK).Value), "")
    wsIE.Cells(6, 3).Value = IIf(colLag > 0, CStr(wsA.Cells(zeile, colLag).Value), "")
    wsIE.Cells(6, 6).Value = IIf(colWG > 0, CStr(wsA.Cells(zeile, colWG).Value), "")
    wsIE.Cells(7, 3).Value = IIf(colAttr > 0, CStr(wsA.Cells(zeile, colAttr).Value), "")
    wsIE.Cells(9, 3).Value = vAnz    ' SOLL
    wsIE.Cells(9, 6).Value = ""      ' GEZAEHLT leeren
    wsIE.Cells(10, 6).Value = ""     ' DIFFERENZ leeren
    wsIE.Cells(10, 6).Interior.ColorIndex = xlNone
    wsIE.Cells(IE_ZEILE_REF, 1).Value = zeile

    ' InvSuche ausblenden
    Dim wsIS As Worksheet: Set wsIS = LagerMakros.GetSheet("InvSuche")
    If Not wsIS Is Nothing Then wsIS.Visible = xlSheetHidden

    wsIE.Visible = xlSheetVisible
    wsIE.Activate
    wsIE.Cells(9, 6).Select
End Sub


' ================================================================
'  InvEingabe - Differenz anzeigen waehrend Eingabe
' ================================================================
Sub InvEingabe_DiffAnzeigen()
    Dim wsIE As Worksheet: Set wsIE = LagerMakros.GetSheet("InvEingabe")
    If wsIE Is Nothing Then Exit Sub

    Dim istStr  As String: istStr = Trim(CStr(wsIE.Cells(9, 6).Value))
    Dim sollStr As String: sollStr = Trim(CStr(wsIE.Cells(9, 3).Value))

    If istStr = "" Then
        wsIE.Cells(10, 6).Value = ""
        wsIE.Cells(10, 6).Interior.ColorIndex = xlNone
        Exit Sub
    End If

    Dim ist  As Double: ist = val(istStr)
    Dim soll As Double: soll = val(sollStr)
    Dim diff As Double: diff = ist - soll

    wsIE.Cells(10, 6).Value = Format(diff, "+0;-0;0")
    If diff = 0 Then
        wsIE.Cells(10, 6).Interior.Color = RGB(198, 239, 206)
        wsIE.Cells(10, 6).Font.Color = RGB(0, 97, 0)
    ElseIf diff > 0 Then
        wsIE.Cells(10, 6).Interior.Color = RGB(255, 235, 156)
        wsIE.Cells(10, 6).Font.Color = RGB(156, 87, 0)
    Else
        wsIE.Cells(10, 6).Interior.Color = RGB(255, 199, 206)
        wsIE.Cells(10, 6).Font.Color = RGB(156, 0, 6)
    End If
End Sub


' ================================================================
'  InvEingabe - Abbrechen -> zurueck zu InvSuche
' ================================================================
Sub InvEingabe_Abbrechen()
    Dim wsIS As Worksheet: Set wsIS = LagerMakros.GetSheet("InvSuche")
    Dim wsIE As Worksheet: Set wsIE = LagerMakros.GetSheet("InvEingabe")
    If Not wsIE Is Nothing Then wsIE.Visible = xlSheetHidden
    If Not wsIS Is Nothing Then
        wsIS.Visible = xlSheetVisible
        wsIS.Activate
    End If
End Sub


' ================================================================
'  InvEingabe - Uebernehmen -> in InvDaten speichern
' ================================================================
Sub InvEingabe_Uebernehmen()
    Dim wsIE As Worksheet: Set wsIE = LagerMakros.GetSheet("InvEingabe")
    If wsIE Is Nothing Then Exit Sub

    Dim istStr As String: istStr = Trim(CStr(wsIE.Cells(9, 6).Value))
    If istStr = "" Then
        MsgBox "Bitte die gezaehlte Menge eingeben.", vbExclamation
        Exit Sub
    End If

    Dim zeile   As Long:   zeile = val(wsIE.Cells(IE_ZEILE_REF, 1).Value)
    Dim ean     As String: ean = CStr(wsIE.Cells(4, 6).Value)
    Dim artNr   As String: artNr = CStr(wsIE.Cells(4, 3).Value)
    Dim artName As String: artName = CStr(wsIE.Cells(3, 3).Value)
    Dim lagerort As String: lagerort = CStr(wsIE.Cells(6, 3).Value)
    Dim soll    As Double: soll = val(CStr(wsIE.Cells(9, 3).Value))
    Dim ist     As Double: ist = val(istStr)

    InvDaten_Speichern ean, artNr, artName, soll, ist, lagerort

    MsgBox "Inventureintrag gespeichert!" & Chr(10) & _
           "Artikel: " & artName & Chr(10) & _
           "SOLL: " & Format(soll, "0") & " / IST: " & Format(ist, "0") & _
           " / Diff: " & Format(ist - soll, "+0;-0;0"), _
           vbInformation

    InvEingabe_Abbrechen
End Sub


' ================================================================
'  Schnellansicht - Oeffnen
' ================================================================
Sub Schnellansicht_Oeffnen()
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If wsS Is Nothing Then
        MsgBox "Schnellansicht-Sheet nicht gefunden!", vbCritical
        Exit Sub
    End If
    Application.ScreenUpdating = False
    wsS.Activate
    LagerMakros.Schnellansicht_Aktualisieren
    Application.ScreenUpdating = True
    Application.StatusBar = "Schnellansicht aktualisiert."
End Sub


' ================================================================
'  Schnellansicht - Schliessen -> Artikel-Sheet
' ================================================================
Sub Schnellansicht_Schliessen()
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Application.StatusBar = False
    If Not wsA Is Nothing Then wsA.Activate
End Sub


' ================================================================
'  Schnellansicht - EK-Preis Spalte ein-/ausblenden
' ================================================================
Sub Schnellansicht_EK_Toggle()
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = wsS.Cells(wsS.Rows.count, 4).End(xlUp).Row
    If lastRow < 4 Then lastRow = 4

    ' EK-Preis ist in Spalte G (7)
    Dim rngEK As Range
    Set rngEK = wsS.Range(wsS.Cells(4, 7), wsS.Cells(lastRow, 7))

    If rngEK.Cells(1, 1).NumberFormat = ";;;" Then
        rngEK.NumberFormat = "#,##0.00 " & ChrW(8364)
        wsS.Cells(3, 7).Value = "EK-Preis " & ChrW(8364)
        wsS.Cells(2, 7).Value = "EK AUSBL."
    Else
        rngEK.NumberFormat = ";;;"
        wsS.Cells(3, 7).Value = ""
        wsS.Cells(2, 7).Value = "EK EINBL."
    End If
End Sub

' ================================================================
'  ARTIKEL TOOLBAR EINRICHTEN (Shapes + Events ins Artikel-Sheet)
'  Erzeugt: Titelzeile (1), Suchzeile (2), 7 Toolbar-Buttons (3)
'  Buttons: GITHUB, NEUER ARTIKEL, ZU-/ABGANG, ETIKETT,
'           EK ausbl., FILTER LOESCHEN, SCHNELLANSICHT
' ================================================================
Sub Setup_Artikel_Toolbar()
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsA Is Nothing Then MsgBox "Artikel-Sheet nicht gefunden!", vbCritical: Exit Sub

    Application.ScreenUpdating = False

    ' 2 Zeilen einfuegen falls noch altes Layout (Header in Zeile 2, Daten ab Zeile 3)
    If InStr(wsA.Cells(1, 1).Value, "Artikel - Lagerverwaltung") = 0 Then
        wsA.Rows("1:2").Insert Shift:=xlShiftDown
    End If

    ' Alte Shapes + alten Inhalt in Zeilen 1-3 entfernen
    Dim shp As Shape
    For Each shp In wsA.Shapes
        On Error Resume Next
        If shp.TopLeftCell.Row <= 3 Then shp.Delete
        On Error GoTo 0
    Next shp
    wsA.Rows("1:3").ClearContents
    wsA.Rows("1:3").Interior.ColorIndex = xlNone
    wsA.Rows("5:5000").Hidden = False

    ' Farben
    Dim blauDunkel As Long: blauDunkel = RGB(32, 55, 100)
    Dim gruen As Long:      gruen = RGB(55, 94, 50)
    Dim orange As Long:     orange = RGB(180, 90, 0)
    Dim grau As Long:       grau = RGB(89, 89, 89)
    Dim blauMittel As Long: blauMittel = RGB(46, 80, 144)
    Dim gitHub As Long:     gitHub = RGB(36, 41, 46)
    Dim stahlBlau As Long:  stahlBlau = RGB(31, 97, 141)
    Dim kopfgrau As Long:   kopfgrau = RGB(64, 64, 64)

    ' ZEILE 1: Titel
    wsA.Rows(1).RowHeight = 28
    On Error Resume Next: wsA.Range("A1:U1").Merge: On Error GoTo 0
    wsA.Cells(1, 1).Value = "Artikel - Lagerverwaltung"
    wsA.Cells(1, 1).Interior.Color = blauDunkel
    wsA.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsA.Cells(1, 1).Font.Bold = True
    wsA.Cells(1, 1).Font.Size = 14
    wsA.Cells(1, 1).HorizontalAlignment = xlCenter

    ' ZEILE 2: Suchfeld + Buttons SUCHEN / LEEREN / AKTUALISIEREN
    wsA.Rows(2).RowHeight = 30
    On Error Resume Next: wsA.Rows(2).UnMerge: On Error GoTo 0
    wsA.Rows(2).Interior.ColorIndex = xlNone
    wsA.Cells(2, 1).Value = "Suche:"
    wsA.Cells(2, 1).Font.Bold = True
    wsA.Cells(2, 1).Font.Size = 14
    wsA.Cells(2, 1).Interior.Color = RGB(242, 242, 242)
    wsA.Cells(2, 1).HorizontalAlignment = xlRight
    On Error Resume Next: wsA.Range("B2:D2").Merge: On Error GoTo 0
    wsA.Cells(2, 2).NumberFormat = "@"
    wsA.Cells(2, 2).Interior.Color = RGB(255, 255, 153)
    wsA.Cells(2, 2).Value = ""
    wsA.Cells(2, 2).Font.Size = 14
    wsA.Cells(2, 2).Font.Color = RGB(0, 0, 0)

    Dim xPos2  As Single: xPos2 = wsA.Range("B2:D2").Left + wsA.Range("B2:D2").Width + 3
    Dim yPos2  As Single: yPos2 = wsA.Rows(2).Top + 1
    Dim btnH2  As Single: btnH2 = wsA.Rows(2).Height - 2
    Dim btnW2  As Single: btnW2 = 75
    Dim wTreffer As Single: wTreffer = 90

    Dim oSuch As Shape
    Set oSuch = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos2, yPos2, btnW2, btnH2)
    With oSuch
        .Name = "btnSuchen": .Fill.ForeColor.RGB = gruen: .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "SUCHEN": .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255): .TextFrame.Characters.Font.Size = 13
        .TextFrame.HorizontalAlignment = xlCenter: .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "LagerMakros.Artikel_Suchen"
    End With
    xPos2 = xPos2 + btnW2 + 4

    Dim oTreffer As Shape
    Set oTreffer = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos2, yPos2, wTreffer, btnH2)
    With oTreffer
        .Name = "trefferAnzeige": .Fill.ForeColor.RGB = RGB(242, 242, 242)
        .Line.ForeColor.RGB = RGB(180, 180, 180): .Line.Visible = msoTrue
        .TextFrame.Characters.Text = "Treffer": .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12: .TextFrame.Characters.Font.Color = RGB(89, 89, 89)
        .TextFrame.HorizontalAlignment = xlCenter: .TextFrame.VerticalAlignment = xlCenter
    End With
    xPos2 = xPos2 + wTreffer + 4

    Dim oLeer As Shape
    Set oLeer = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos2, yPos2, btnW2, btnH2)
    With oLeer
        .Name = "btnLeeren": .Fill.ForeColor.RGB = orange: .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "LEEREN": .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255): .TextFrame.Characters.Font.Size = 13
        .TextFrame.HorizontalAlignment = xlCenter: .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "LagerMakros.Artikel_Suche_Leeren"
    End With
    xPos2 = xPos2 + btnW2 + 4

    Dim oAkt As Shape
    Set oAkt = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos2, yPos2, btnW2 + 30, btnH2)
    With oAkt
        .Name = "btnAktualisieren": .Fill.ForeColor.RGB = RGB(0, 112, 96): .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "AKTUALISIEREN": .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255): .TextFrame.Characters.Font.Size = 11
        .TextFrame.HorizontalAlignment = xlCenter: .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "LagerMakros.Artikel_Aktualisieren"
    End With

    ' ZEILE 3: 7 Toolbar-Buttons
    wsA.Rows(3).RowHeight = 28
    Dim btnH3 As Single: btnH3 = wsA.Rows(3).Height - 4
    Dim btnW3 As Single: btnW3 = 108
    Dim gap3  As Single: gap3 = 5
    Dim xStart As Single: xStart = wsA.Cells(3, 1).Left + 4

    Dim btns As Variant
    btns = Array( _
        Array("GITHUB", gitHub, "NeueModule.GitHub_Export"), _
        Array("NEUER ARTIKEL", stahlBlau, "LagerMakros.NeuerArtikel"), _
        Array("ZU-/ABGANG", grau, "LagerMakros.ZuAbgang_Buchen"), _
        Array("ETIKETT", grau, "LagerMakros.Etikett_Drucken"), _
        Array("EK ausbl.", blauMittel, "LagerMakros.EK_Toggle"), _
        Array("FILTER LOESCHEN", orange, "LagerMakros.Filter_Loeschen"), _
        Array("SCHNELLANSICHT", blauMittel, "NeueModule.Schnellansicht_Oeffnen"))

    Dim xPos As Single: xPos = xStart
    Dim yPos As Single: yPos = wsA.Cells(3, 1).Top + 2
    Dim bi As Integer
    For bi = 0 To 6
        Dim oB As Shape
        Set oB = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos, yPos, btnW3, btnH3)
        With oB
            .Name = "btnT3_" & bi: .Fill.ForeColor.RGB = btns(bi)(1): .Line.Visible = msoFalse
            .TextFrame.Characters.Text = btns(bi)(0): .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255): .TextFrame.Characters.Font.Size = 10
            .TextFrame.HorizontalAlignment = xlCenter: .TextFrame.VerticalAlignment = xlCenter
            .OnAction = btns(bi)(2)
        End With
        xPos = xPos + btnW3 + gap3
    Next bi

    ' ZEILE 4: Spaltenkoepfe
    wsA.Rows(4).RowHeight = 22
    With wsA.Range("A4:U4")
        .Interior.Color = kopfgrau: .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True: .Font.Size = 10: .HorizontalAlignment = xlCenter
    End With

    ' Zeilen 1-4 einfrieren
    wsA.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    wsA.Cells(5, 1).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' Event-Code einbetten
    Dim vbA As Object: Set vbA = ThisWorkbook.VBProject.VBComponents(wsA.CodeName)
    Dim cmA As Object: Set cmA = vbA.CodeModule
    If cmA.CountOfLines > 0 Then cmA.DeleteLines 1, cmA.CountOfLines
    Dim ev As String: ev = ""
    ev = ev & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    ev = ev & "    If Target.Row >= 5 Then" & Chr(10)
    ev = ev & "        Cancel = True" & Chr(10)
    ev = ev & "        NeueModule.ArtikelDetail_Laden Target.Row" & Chr(10)
    ev = ev & "    End If" & Chr(10)
    ev = ev & "End Sub" & Chr(10) & Chr(10)
    ev = ev & "Private Sub Worksheet_Activate()" & Chr(10)
    ev = ev & "    On Error Resume Next" & Chr(10)
    ev = ev & "    Me.Cells(2, 2).Select" & Chr(10)
    ev = ev & "End Sub" & Chr(10) & Chr(10)
    ev = ev & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    ev = ev & "    If Not Intersect(Target, Me.Range(""B2"")) Is Nothing Then" & Chr(10)
    ev = ev & "        On Error GoTo Fertig" & Chr(10)
    ev = ev & "        Application.EnableEvents = False" & Chr(10)
    ev = ev & "        LagerMakros.Artikel_Suchen" & Chr(10)
    ev = ev & "        Application.EnableEvents = True" & Chr(10)
    ev = ev & "    End If" & Chr(10)
    ev = ev & "    Exit Sub" & Chr(10)
    ev = ev & "Fertig: Application.EnableEvents = True" & Chr(10)
    ev = ev & "End Sub" & Chr(10) & Chr(10)
    ev = ev & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10)
    ev = ev & "    If Target.Row >= 5 Then LagerMakros.Artikel_Zeile_Markieren Target" & Chr(10)
    ev = ev & "End Sub" & Chr(10)
    cmA.AddFromString ev

    Application.ScreenUpdating = True
    Application.StatusBar = "Artikel Toolbar eingerichtet."
End Sub

' ================================================================
'  GITHUB EXPORT
'  Exportiert .bas-Module + lager.json und pusht per git
' ================================================================
Sub GitHub_Export()
    On Error GoTo Fehler

    Const GIT_DIR     As String = "D:\_KI-Projekte-2026\07_Lagerverwaltung-Excel\Lagerverw. FB300\Lagerverwaltung_2026_V3\"
    Const LAGER_SHEET As String = "Schnellansicht"

    ' VBA-Module exportieren
    Dim moduleNames As Variant
    moduleNames = Array("LagerMakros", "NeueModule", "NeuArtikelModul", "ArtikelFix", "DuplikatPruefer")

    Dim vbComp  As Object
    Dim modName As String
    Dim j       As Integer
    For j = 0 To UBound(moduleNames)
        modName = moduleNames(j)
        For Each vbComp In ThisWorkbook.VBProject.VBComponents
            If vbComp.Name = modName Then
                vbComp.Export GIT_DIR & modName & ".bas"
                Exit For
            End If
        Next vbComp
    Next j

    ' Lagerdaten als JSON exportieren
    Call ExportLagerJSON(GIT_DIR, LAGER_SHEET)

    ' Git: add -> commit -> push
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim ts As String: ts = Format(Now(), "DD.MM.YYYY HH\:MM")
    Dim commitMsg As String: commitMsg = "Lagerverwaltung Update - " & ts
    Dim gitCmd As String
    gitCmd = "cmd /c cd /d """ & GIT_DIR & """" & _
             " && git add -A" & _
             " && git commit -m """ & commitMsg & """" & _
             " && git push"
    Dim rc As Long: rc = sh.Run(gitCmd, 1, True)

    If rc = 0 Then
        MsgBox "Erfolgreich nach GitHub gepusht!" & vbNewLine & "Commit: " & commitMsg, _
               vbInformation, "GitHub Export"
    Else
        MsgBox "Git meldete Fehler-Code " & rc & "." & vbNewLine & _
               "Bitte pruefe das CMD-Fenster.", vbExclamation, "GitHub Export - Warnung"
    End If
    Exit Sub

Fehler:
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, vbCritical, "GitHub Export - Fehler"
End Sub

Private Sub ExportLagerJSON(targetDir As String, sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Const C_NR          As Integer = 2
    Const C_ARTNR       As Integer = 3
    Const C_ARTIKEL     As Integer = 4
    Const C_EAN         As Integer = 5
    Const C_VK          As Integer = 6
    Const C_EK          As Integer = 7
    Const C_BESTAND     As Integer = 8
    Const C_EINHEIT     As Integer = 9
    Const C_LAGERORT    As Integer = 10
    Const C_WARENGRUPPE As Integer = 11
    Const C_ATTRIBUT    As Integer = 12

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, C_ARTNR).End(xlUp).Row

    Dim json As String: Dim isFirst As Boolean: isFirst = True
    json = "[" & vbNewLine
    Dim i As Long: Dim artNr As String
    For i = 3 To lastRow
        artNr = Trim(CStr(ws.Cells(i, C_ARTNR).Value))
        If artNr = "" Then GoTo nextRow
        If Not isFirst Then json = json & "," & vbNewLine
        isFirst = False
        json = json & "  {" & _
            """nr"":" & JStr(ws.Cells(i, C_NR)) & "," & _
            """artnr"":" & JStr(ws.Cells(i, C_ARTNR)) & "," & _
            """artikel"":" & JStr(ws.Cells(i, C_ARTIKEL)) & "," & _
            """ean"":" & JStr(ws.Cells(i, C_EAN)) & "," & _
            """vk"":" & JStr(ws.Cells(i, C_VK)) & "," & _
            """ek"":" & JStr(ws.Cells(i, C_EK)) & "," & _
            """bestand"":" & JStr(ws.Cells(i, C_BESTAND)) & "," & _
            """einheit"":" & JStr(ws.Cells(i, C_EINHEIT)) & "," & _
            """lagerort"":" & JStr(ws.Cells(i, C_LAGERORT)) & "," & _
            """warengruppe"":" & JStr(ws.Cells(i, C_WARENGRUPPE)) & "," & _
            """attribut"":" & JStr(ws.Cells(i, C_ATTRIBUT)) & "}"
nextRow:
    Next i
    json = json & vbNewLine & "]"

    Dim fNum As Integer: fNum = FreeFile
    Open targetDir & "lager.json" For Output As #fNum
    Print #fNum, json
    Close #fNum
End Sub

Private Function JStr(cell As Range) As String
    Dim s As String: s = CStr(cell.Value)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    JStr = """" & s & """"
End Function
