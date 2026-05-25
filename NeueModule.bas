Attribute VB_Name = "NeueModule"
Public g_AufruferSheet As String   ' Merkt wo ArtikelDetail aufgerufen wurde
Option Explicit

' ================================================================
'  KONSTANTEN fuer neue Module
' ================================================================
Const AD_ZEILE_REF  As Long = 17   ' Zeile in ArtikelDetail die die Artikelzeile speichert
Const IE_ZEILE_REF  As Long = 14   ' Zeile in InvEingabe die die Artikelzeile speichert
Const ID_DATEN_START As Long = 4   ' InvDaten: Datenzeilen ab Zeile 4

' ================================================================
'  SETUP - ArtikelDetail Sheet erstellen / neu aufbauen
' ================================================================
Sub Setup_ArtikelDetail()
    Dim wsAD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "ArtikelDetail" Then Set wsAD = ws: Exit For
    Next ws
    If wsAD Is Nothing Then
        Set wsAD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAD.Name = "ArtikelDetail"
    End If

    Application.ScreenUpdating = False
    wsAD.Cells.Clear
    wsAD.Cells.Interior.ColorIndex = xlNone

    Dim blau As Long:    blau = RGB(32, 55, 100)
    Dim hellblau As Long: hellblau = RGB(46, 80, 144)
    Dim gruen As Long:   gruen = RGB(55, 110, 50)
    Dim orange As Long:  orange = RGB(180, 90, 0)
    Dim gelb As Long:    gelb = RGB(255, 255, 0)
    Dim hellgrau As Long: hellgrau = RGB(242, 242, 242)
    Dim rot As Long:     rot = RGB(192, 0, 0)

    ' Zeile 1: Titel
    wsAD.Range("A1:G1").Merge
    wsAD.Cells(1, 1).Value = "Artikel - Details"
    wsAD.Cells(1, 1).Interior.Color = blau
    wsAD.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsAD.Cells(1, 1).Font.Size = 14
    wsAD.Cells(1, 1).Font.Bold = True
    wsAD.Cells(1, 1).HorizontalAlignment = xlCenter
    wsAD.Rows(1).RowHeight = 30

    ' Felder definieren: Zeile, Beschriftung Spalte B, Wert Spalte C, Beschriftung Spalte E, Wert Spalte F
    Dim felder As Variant
    felder = Array( _
        Array(3, "Artikel:", "", "", ""), _
        Array(4, "Art.-Nr.:", "", "EAN:", ""), _
        Array(5, "VK-Preis:", "", "VK2 (opt.):", ""), _
        Array(6, "EK-Preis:", "", "", ""), _
        Array(7, "MwSt %:", "", "Lieferant:", ""), _
        Array(8, "Bestand:", "", "Einheit:", ""), _
        Array(9, "Lagerort:", "", "Warengruppe:", ""), _
        Array(10, "Attribut:", "", "", ""), _
        Array(11, "TextA:", "", "", ""), _
        Array(12, "TextB:", "", "", ""), _
        Array(13, "Rohgewinn:", "", "", "") _
    )

    Dim i As Integer
    For i = 0 To UBound(felder)
        Dim r As Long: r = felder(i)(0)
        wsAD.Cells(r, 2).Value = felder(i)(1)
        wsAD.Cells(r, 2).Font.Bold = True
        wsAD.Cells(r, 2).Font.Size = 11
        wsAD.Cells(r, 3).Interior.Color = gelb
        wsAD.Cells(r, 3).Font.Size = 11
        wsAD.Range(wsAD.Cells(r, 3), wsAD.Cells(r, 4)).Merge
        If felder(i)(3) <> "" Then
            wsAD.Cells(r, 5).Value = felder(i)(3)
            wsAD.Cells(r, 5).Font.Bold = True
            wsAD.Cells(r, 5).Font.Size = 11
            wsAD.Cells(r, 6).Interior.Color = gelb
            wsAD.Cells(r, 6).Font.Size = 11
            wsAD.Range(wsAD.Cells(r, 6), wsAD.Cells(r, 7)).Merge
        End If
        wsAD.Rows(r).RowHeight = 22
    Next i

    ' Zeile 3: Artikel-Name breiter
    wsAD.Range("C3:G3").Merge
    wsAD.Cells(3, 3).Interior.Color = gelb
    wsAD.Rows(3).RowHeight = 26

    ' Zeile 13: Rohgewinn - schreibgeschuetzt / berechnet
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
    wsAD.Range("B15:C15").Merge
    wsAD.Rows(15).RowHeight = 28

    With wsAD.Cells(15, 4)
        .Value = "BEWEGUNGEN"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = orange: .HorizontalAlignment = xlCenter
    End With
    wsAD.Range("D15:E15").Merge

    With wsAD.Cells(15, 6)
        .Value = "SPEICHERN"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = gruen: .HorizontalAlignment = xlCenter
    End With
    wsAD.Range("F15:G15").Merge

    ' Zeile 16: Leerzeile
    wsAD.Rows(16).RowHeight = 6

    ' Zeile 17: versteckt - speichert Artikelzeile
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

    ' Dropdown-Listen fuer Lagerort, Warengruppe, Attribut
    Dim wsLag As Worksheet: Set wsLag = LagerMakros.GetSheet("Lagerorte")
    Dim wsWG  As Worksheet: Set wsWG = LagerMakros.GetSheet("Warengruppen")
    Dim wsA2  As Worksheet: Set wsA2 = LagerMakros.GetSheet("Artikel")

    ' Lagerort (Zeile 9, Spalte 3)
    On Error Resume Next
    wsAD.Cells(9, 3).Validation.Delete
    If Not wsLag Is Nothing Then
        Dim lastLag As Long: lastLag = wsLag.Cells(wsLag.Rows.Count, 1).End(xlUp).Row
        If lastLag >= 2 Then
            wsAD.Cells(9, 3).Validation.Add Type:=xlValidateList, _
                Formula1:="=Lagerorte!$A$2:$A$" & lastLag
        End If
    End If

    ' Warengruppe (Zeile 9, Spalte 6)
    wsAD.Cells(9, 6).Validation.Delete
    If Not wsWG Is Nothing Then
        Dim lastWG As Long: lastWG = wsWG.Cells(wsWG.Rows.Count, 1).End(xlUp).Row
        If lastWG >= 2 Then
            wsAD.Cells(9, 6).Validation.Add Type:=xlValidateList, _
                Formula1:="=Warengruppen!$A$2:$A$" & lastWG
        End If
    End If

    ' Attribut (Zeile 10, Spalte 3) - aus Attribut-Spalte im Artikel-Sheet
    wsAD.Cells(10, 3).Validation.Delete
    If Not wsA2 Is Nothing Then
        Dim colAttr2 As Long: colAttr2 = LagerMakros.Spalte_Finden(wsA2, "Attribut")
        If colAttr2 > 0 Then
            Dim lastA2 As Long: lastA2 = wsA2.Cells(wsA2.Rows.Count, colAttr2).End(xlUp).Row
            If lastA2 >= 3 Then
                wsAD.Cells(10, 3).Validation.Add Type:=xlValidateList, _
                    Formula1:="=Artikel!$" & Chr(64 + colAttr2) & "$3:$" & Chr(64 + colAttr2) & "$" & lastA2
            End If
        End If
    End If
    On Error GoTo 0

    ' Sheet-Events installieren
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsAD.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim c As String: c = ""
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    Cancel = True" & Chr(10)
    c = c & "    Dim col As Long: col = Target.Column" & Chr(10)
    c = c & "    Dim rw As Long:  rw  = Target.Row" & Chr(10)
    c = c & "    If rw = 15 And col >= 2 And col <= 3 Then NeueModule.ArtikelDetail_Schliessen" & Chr(10)
    c = c & "    If rw = 15 And col >= 4 And col <= 5 Then NeueModule.ArtikelDetail_Bewegungen" & Chr(10)
    c = c & "    If rw = 15 And col >= 6 And col <= 7 Then NeueModule.ArtikelDetail_Speichern" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c

    Application.ScreenUpdating = True
End Sub

' ================================================================
'  ArtikelDetail - Artikel laden (aus Artikel-Sheet Zeile)
' ================================================================
Sub ArtikelDetail_Laden(zeile As Long)
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    If wsA Is Nothing Or wsAD Is Nothing Then
        MsgBox "ArtikelDetail-Blatt fehlt. Bitte Setup_ArtikelDetail ausfuehren.", vbExclamation
        Exit Sub
    End If

    ' Aufrufer merken (BEVOR irgendwas versteckt wird)
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If Not wsS Is Nothing And ActiveSheet.Name = wsS.Name Then
        g_AufruferSheet = "Schnellansicht"
    Else
        g_AufruferSheet = "Artikel"
    End If

    On Error GoTo ArtikelDetailFehler

    Dim colArt   As Long: colArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr    As Long: colNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim colEAN   As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colVK    As Long: colVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK    As Long: colEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim colMwst  As Long: colMwst = LagerMakros.Spalte_Finden(wsA, "MWST")
    Dim colLief  As Long: colLief = LagerMakros.Spalte_Finden(wsA, "LIEFERANT")
    Dim colAnz   As Long: colAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim colEinh  As Long: colEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim colLag   As Long: colLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim colWG    As Long: colWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colAttr  As Long: colAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")
    Dim colTA    As Long: colTA = LagerMakros.Spalte_Finden(wsA, "TextA")
    Dim colTB    As Long: colTB = LagerMakros.Spalte_Finden(wsA, "TextB")
    Dim colVK2   As Long: colVK2 = LagerMakros.Spalte_Finden(wsA, "VK2")

    ' Werte sicher holen (kein IIf mit Cells-Referenz - IIf wertet beide Seiten aus!)
    Dim vArt  As String:  If colArt > 0 Then vArt = CStr(wsA.Cells(zeile, colArt).Value)
    Dim vNr   As String:  If colNr > 0 Then vNr = CStr(wsA.Cells(zeile, colNr).Value)
    Dim vLief As String:  If colLief > 0 Then vLief = CStr(wsA.Cells(zeile, colLief).Value)
    Dim vEinh As String:  If colEinh > 0 Then vEinh = CStr(wsA.Cells(zeile, colEinh).Value) Else vEinh = "Stk"
    Dim vLag  As String:  If colLag > 0 Then vLag = CStr(wsA.Cells(zeile, colLag).Value)
    Dim vWG   As String:  If colWG > 0 Then vWG = CStr(wsA.Cells(zeile, colWG).Value)
    Dim vAttr As String:  If colAttr > 0 Then vAttr = CStr(wsA.Cells(zeile, colAttr).Value)
    Dim vTA   As String:  If colTA > 0 Then vTA = CStr(wsA.Cells(zeile, colTA).Value)
    Dim vTB   As String:  If colTB > 0 Then vTB = CStr(wsA.Cells(zeile, colTB).Value)
    Dim vVK   As Double:  If colVK > 0 Then vVK = Val(wsA.Cells(zeile, colVK).Value)
    Dim vEK   As Double:  If colEK > 0 Then vEK = Val(wsA.Cells(zeile, colEK).Value)
    Dim vMwst As Double:  If colMwst > 0 Then vMwst = Val(wsA.Cells(zeile, colMwst).Value) Else vMwst = 19
    Dim vAnz  As Double:  If colAnz > 0 Then vAnz = Val(wsA.Cells(zeile, colAnz).Value)
    Dim vVK2  As Double:  If colVK2 > 0 Then vVK2 = Val(wsA.Cells(zeile, colVK2).Value)

    Dim vEAN  As String
    If colEAN > 0 Then
        Dim rawEAN As Variant: rawEAN = wsA.Cells(zeile, colEAN).Value
        If IsNumeric(rawEAN) Then
            vEAN = Format(CDbl(rawEAN), "0000000000000")
        Else
            vEAN = CStr(rawEAN)
        End If
    End If

    Dim rohgewinn As Double: rohgewinn = vVK - vEK
    Dim marge     As Double: If vVK > 0 Then marge = rohgewinn / vVK * 100
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

    ' Schnellansicht jetzt erst ausblenden, ArtikelDetail zeigen
    If g_AufruferSheet = "Schnellansicht" Then
        If Not wsS Is Nothing Then wsS.Visible = xlSheetHidden
    End If
    wsAD.Visible = xlSheetVisible
    wsAD.Activate
    On Error Resume Next
    wsAD.Cells(3, 3).Select
    On Error GoTo 0
    Exit Sub

ArtikelDetailFehler:
    MsgBox "Fehler " & Err.Number & " in ArtikelDetail_Laden: " & Err.Description, vbCritical
    g_AufruferSheet = ""
End Sub


' ================================================================
'  ArtikelDetail - Laden via EAN oder Artikelname (aus Schnellansicht)
' ================================================================
Sub ArtikelDetail_LadenNachEAN(sEAN As String, sArtName As String)
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub

    Dim colEAN As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colArt As Long: colArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    If colArt = 0 Then Exit Sub

    Dim lastRow As Long: lastRow = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
    Dim gefunden As Long: gefunden = 0
    Dim i As Long

    ' Zuerst per EAN suchen
    If sEAN <> "" And colEAN > 0 Then
        For i = 5 To lastRow
            If CStr(wsA.Cells(i, colEAN).Value) = sEAN Then
                gefunden = i: Exit For
            End If
        Next i
    End If

    ' Falls EAN nicht gefunden: per Artikelname suchen
    If gefunden = 0 And sArtName <> "" Then
        For i = 5 To lastRow
            If CStr(wsA.Cells(i, colArt).Value) = sArtName Then
                gefunden = i: Exit For
            End If
        Next i
    End If

    If gefunden > 0 Then
        ArtikelDetail_Laden gefunden
    Else
        MsgBox "Artikel nicht gefunden: " & sArtName, vbExclamation
    End If
End Sub

' ================================================================
'  ArtikelDetail - Aenderungen zurueck ins Artikel-Sheet speichern
' ================================================================
Sub ArtikelDetail_Speichern()
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    If wsA Is Nothing Or wsAD Is Nothing Then Exit Sub

    Dim zeile As Long: zeile = Val(wsAD.Cells(AD_ZEILE_REF, 1).Value)
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
    If colVK > 0 Then wsA.Cells(zeile, colVK).Value = Val(wsAD.Cells(5, 3).Value)
    If colEK > 0 Then wsA.Cells(zeile, colEK).Value = Val(wsAD.Cells(6, 3).Value)
    If colMwst > 0 Then wsA.Cells(zeile, colMwst).Value = Val(wsAD.Cells(7, 3).Value)
    If colLief > 0 Then wsA.Cells(zeile, colLief).Value = wsAD.Cells(7, 6).Value
    If colAnz > 0 Then wsA.Cells(zeile, colAnz).Value = Val(wsAD.Cells(8, 3).Value)
    If colEinh > 0 Then wsA.Cells(zeile, colEinh).Value = wsAD.Cells(8, 6).Value
    If colLag > 0 Then wsA.Cells(zeile, colLag).Value = wsAD.Cells(9, 3).Value
    If colWG > 0 Then wsA.Cells(zeile, colWG).Value = wsAD.Cells(9, 6).Value
    If colAttr > 0 Then wsA.Cells(zeile, colAttr).Value = wsAD.Cells(10, 3).Value
    If colTA > 0 Then wsA.Cells(zeile, colTA).Value = wsAD.Cells(11, 3).Value
    If colTB > 0 Then wsA.Cells(zeile, colTB).Value = wsAD.Cells(12, 3).Value
    If colVK2 > 0 Then wsA.Cells(zeile, colVK2).Value = Val(wsAD.Cells(5, 6).Value)

    ' Schnellansicht aktualisieren
    On Error Resume Next
    LagerMakros.Schnellansicht_Aktualisieren
    On Error GoTo 0

    Application.StatusBar = "Artikel gespeichert."
    ArtikelDetail_Schliessen
End Sub

' ================================================================
'  ArtikelDetail - Schliessen -> zurueck zum Artikel-Sheet
' ================================================================
Sub ArtikelDetail_Schliessen()
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsAD As Worksheet: Set wsAD = LagerMakros.GetSheet("ArtikelDetail")
    If wsA Is Nothing Then Exit Sub
    Dim zeile As Long
    If Not wsAD Is Nothing Then zeile = Val(wsAD.Cells(AD_ZEILE_REF, 1).Value)
    Application.StatusBar = False
    If Not wsAD Is Nothing Then wsAD.Visible = xlSheetHidden
    If g_AufruferSheet = "Schnellansicht" Then
        Dim wsSrck As Worksheet: Set wsSrck = LagerMakros.GetSheet("Schnell")
        If Not wsSrck Is Nothing Then
            wsSrck.Visible = xlSheetVisible
            wsSrck.Activate
        End If
    Else
        wsA.Activate
        If zeile >= 3 Then Application.GoTo wsA.Cells(zeile, 1), True
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
    Dim artName As String: artName = wsAD.Cells(3, 3).Value
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
        Set wsBP = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsBP.Name = "BewPopup"
    End If

    Application.ScreenUpdating = False
    wsBP.Cells.Clear

    Dim blau As Long: blau = RGB(31, 56, 100)

    ' Zeile 1: Titel
    wsBP.Range("A1:F1").Merge
    wsBP.Cells(1, 1).Value = "Bewegungshistorie"
    wsBP.Cells(1, 1).Interior.Color = blau
    wsBP.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsBP.Cells(1, 1).Font.Bold = True
    wsBP.Cells(1, 1).Font.Size = 13
    wsBP.Cells(1, 1).HorizontalAlignment = xlCenter
    wsBP.Rows(1).RowHeight = 28

    ' Zeile 2: Artikelname + Schliessen-Button
    wsBP.Range("B2:E2").Merge
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
    hdrs = Array("Datum/Zeit", "", "Menge", "Typ", "Lagerort", "Benutzer")
    Dim j As Integer
    For j = 0 To 5
        wsBP.Cells(3, j + 1).Value = hdrs(j)
        wsBP.Cells(3, j + 1).Interior.Color = RGB(46, 80, 144)
        wsBP.Cells(3, j + 1).Font.Color = RGB(255, 255, 255)
        wsBP.Cells(3, j + 1).Font.Bold = True
    Next j
    wsBP.Rows(3).RowHeight = 20

    ' Spaltenbreiten
    wsBP.Columns(1).ColumnWidth = 8
    wsBP.Columns(2).ColumnWidth = 22
    wsBP.Columns(3).ColumnWidth = 10
    wsBP.Columns(4).ColumnWidth = 10
    wsBP.Columns(5).ColumnWidth = 16
    wsBP.Columns(6).ColumnWidth = 12

    ' Sheet-Events
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsBP.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim c As String: c = ""
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    If Target.Row = 2 And Target.Column = 6 Then" & Chr(10)
    c = c & "        Cancel = True" & Chr(10)
    c = c & "        NeueModule.BewPopup_Schliessen" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c

    Application.ScreenUpdating = True
End Sub

' ================================================================
'  BewPopup - Bewegungshistorie laden
' ================================================================
Sub BewPopup_Laden(ean As String, artName As String)
    Dim wsBP As Worksheet: Set wsBP = LagerMakros.GetSheet("BewPopup")
    Dim wsZ  As Worksheet: Set wsZ = LagerMakros.GetSheet("Abg")
    If wsBP Is Nothing Then
        MsgBox "BewPopup-Sheet fehlt. Bitte Setup_BewPopup ausfuehren.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Alte Daten loeschen
    Dim lastRow As Long: lastRow = wsBP.Cells(wsBP.Rows.Count, 2).End(xlUp).Row
    If lastRow >= 4 Then wsBP.Range("A4:F" & lastRow).Clear

    ' Artikelname in Zeile 2
    wsBP.Range("B2:E2").Value = Left(artName, 40) & "  [EAN: " & ean & "]"

    ' Bewegungen suchen
    Dim treffer As Long: treffer = 0
    If Not wsZ Is Nothing Then
        Dim colEAN_Z As Long: colEAN_Z = LagerMakros.Spalte_Finden(wsZ, "EAN13")
        Dim colDat   As Long: colDat = 1    ' Spalte 1 = Datum/Zeit
        Dim colMenge As Long: colMenge = 5  ' Spalte 5 = Menge
        Dim colTyp   As Long: colTyp = 6    ' Spalte 6 = Typ
        Dim colLag   As Long: colLag = 7    ' Spalte 7 = Lagerort
        Dim colBen   As Long: colBen = 8    ' Spalte 8 = Benutzer

        Dim lastZ As Long: lastZ = wsZ.Cells(wsZ.Rows.Count, 1).End(xlUp).Row
        Dim sRow As Long: sRow = 4
        Dim i As Long
        For i = 2 To lastZ
            If colEAN_Z > 0 Then
                If CStr(wsZ.Cells(i, colEAN_Z).Value) = ean Then
                    wsBP.Cells(sRow, 1).Value = wsZ.Cells(i, colDat).Value
                    wsBP.Cells(sRow, 1).NumberFormat = "DD.MM.YY HH:MM"
                    wsBP.Cells(sRow, 3).Value = wsZ.Cells(i, colMenge).Value
                    wsBP.Cells(sRow, 4).Value = wsZ.Cells(i, colTyp).Value
                    wsBP.Cells(sRow, 5).Value = wsZ.Cells(i, colLag).Value
                    wsBP.Cells(sRow, 6).Value = wsZ.Cells(i, colBen).Value
                    If sRow Mod 2 = 0 Then wsBP.Range("A" & sRow & ":F" & sRow).Interior.Color = RGB(242, 242, 242)
                    sRow = sRow + 1
                    treffer = treffer + 1
                End If
            End If
        Next i
    End If

    If treffer = 0 Then
        wsBP.Cells(4, 2).Value = "(keine Bewegungen gefunden)"
    End If

    Application.ScreenUpdating = True
    wsBP.Visible = xlSheetVisible
    wsBP.Activate
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
        LagerMakros.GetSheet("Artikel").Activate
    End If
    If Not wsBP Is Nothing Then wsBP.Visible = xlSheetHidden
End Sub

' ================================================================
'  SETUP - InvDaten Sheet erstellen (Persistenzspeicher)
' ================================================================
Sub Setup_InvDaten()
    Dim wsID As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "InvDaten" Then Set wsID = ws: Exit For
    Next ws
    If wsID Is Nothing Then
        Set wsID = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsID.Name = "InvDaten"
    End If

    ' Nur Header setzen falls leer
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
'  InvDaten - Eintrag speichern
' ================================================================
Sub InvDaten_Speichern(ean As String, artNr As String, artName As String, _
                       soll As Double, ist As Double, lagerort As String)
    Dim wsID As Worksheet: Set wsID = LagerMakros.GetSheet("InvDaten")
    If wsID Is Nothing Then
        MsgBox "InvDaten-Sheet fehlt. Bitte Setup_InvDaten ausfuehren.", vbExclamation
        Exit Sub
    End If

    ' Pruefen ob Eintrag schon vorhanden (EAN in Spalte 2)
    Dim lastID As Long: lastID = wsID.Cells(wsID.Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    For i = ID_DATEN_START To lastID
        If CStr(wsID.Cells(i, 2).Value) = ean Then
            ' Update bestehenden Eintrag
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
'  InvDaten - IST-Wert fuer EAN laden (sitzungsuebergreifend)
' ================================================================
Function InvDaten_IstLaden(ean As String) As Double
    InvDaten_IstLaden = -1  ' -1 = nicht gefunden
    Dim wsID As Worksheet: Set wsID = LagerMakros.GetSheet("InvDaten")
    If wsID Is Nothing Then Exit Function
    Dim lastID As Long: lastID = wsID.Cells(wsID.Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    For i = ID_DATEN_START To lastID
        If CStr(wsID.Cells(i, 2).Value) = ean Then
            InvDaten_IstLaden = Val(wsID.Cells(i, 6).Value)
            Exit Function
        End If
    Next i
End Function

' ================================================================
'  InvDaten - Inventur abschliessen (Bestaende aktualisieren)
' ================================================================
Sub InvDaten_Abschliessen()
    Dim wsID As Worksheet: Set wsID = LagerMakros.GetSheet("InvDaten")
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsZ  As Worksheet: Set wsZ = LagerMakros.GetSheet("Abg")
    If wsID Is Nothing Or wsA Is Nothing Then
        MsgBox "InvDaten-Sheet oder Artikel-Sheet nicht gefunden!", vbCritical
        Exit Sub
    End If

    If MsgBox("Inventur abschliessen?" & Chr(10) & Chr(10) & _
              "Alle Artikel mit Differenz <> 0 werden im Bewegungsblatt eingetragen" & Chr(10) & _
              "und der Bestand im Artikel-Sheet wird korrigiert.", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Dim colEAN  As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colAnz  As Long: colAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    If colEAN = 0 Or colAnz = 0 Then
        MsgBox "Spalten EAN13 oder ANZAHL nicht gefunden!", vbCritical
        Exit Sub
    End If

    Dim lastID As Long: lastID = wsID.Cells(wsID.Rows.Count, 2).End(xlUp).Row
    Dim lastA  As Long: lastA = wsA.Cells(wsA.Rows.Count, colEAN).End(xlUp).Row
    Dim updated As Long: updated = 0
    Dim i As Long, j As Long

    Application.ScreenUpdating = False

    For i = ID_DATEN_START To lastID
        Dim diff As Double: diff = Val(wsID.Cells(i, 7).Value)
        Dim ist  As Double: ist = Val(wsID.Cells(i, 6).Value)
        Dim eanI As String: eanI = CStr(wsID.Cells(i, 2).Value)
        If eanI = "" Then GoTo WeiterID

        ' Bestand aktualisieren
        For j = 5 To lastA
            If CStr(wsA.Cells(j, colEAN).Value) = eanI Then
                wsA.Cells(j, colAnz).Value = ist
                updated = updated + 1

                ' Abweichungen im Bewegungsblatt eintragen
                If diff <> 0 And Not wsZ Is Nothing Then
                    Dim nRow As Long
                    nRow = wsZ.Cells(wsZ.Rows.Count, 1).End(xlUp).Row + 1
                    wsZ.Cells(nRow, 1).Value = Now()
                    wsZ.Cells(nRow, 1).NumberFormat = "DD.MM.YYYY HH:MM"
                    wsZ.Cells(nRow, 2).Value = eanI
                    wsZ.Cells(nRow, 3).Value = wsID.Cells(i, 3).Value
                    wsZ.Cells(nRow, 4).Value = wsID.Cells(i, 4).Value
                    wsZ.Cells(nRow, 5).Value = Abs(diff)
                    wsZ.Cells(nRow, 6).Value = IIf(diff > 0, "Zugang (Inventur)", "Abgang (Inventur)")
                    wsZ.Cells(nRow, 7).Value = wsID.Cells(i, 8).Value
                    wsZ.Cells(nRow, 8).Value = BENUTZER
                End If
                Exit For
            End If
        Next j
WeiterID:
    Next i

    Application.ScreenUpdating = True
    On Error Resume Next
    LagerMakros.Schnellansicht_Aktualisieren
    On Error GoTo 0

    MsgBox updated & " Bestaende aktualisiert." & Chr(10) & _
           "Bewegungsblatt wurde eingetragen.", vbInformation
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
        Set wsIE = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsIE.Name = "InvEingabe"
    End If

    Application.ScreenUpdating = False
    wsIE.Cells.Clear
    wsIE.Cells.Interior.ColorIndex = xlNone

    Dim blau As Long:    blau = RGB(32, 55, 100)
    Dim hellblau As Long: hellblau = RGB(46, 80, 144)
    Dim gruen As Long:   gruen = RGB(55, 110, 50)
    Dim rot As Long:     rot = RGB(192, 0, 0)
    Dim gelb As Long:    gelb = RGB(255, 255, 200)
    Dim hellgrau As Long: hellgrau = RGB(242, 242, 242)

    ' Zeile 1: Titel
    wsIE.Range("A1:F1").Merge
    wsIE.Cells(1, 1).Value = "Inventur - Artikel pruefen"
    wsIE.Cells(1, 1).Interior.Color = blau
    wsIE.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsIE.Cells(1, 1).Font.Size = 14
    wsIE.Cells(1, 1).Font.Bold = True
    wsIE.Cells(1, 1).HorizontalAlignment = xlCenter
    wsIE.Rows(1).RowHeight = 30

    ' Felder aufbauen
    Dim felder2 As Variant
    felder2 = Array( _
        Array(3, "Artikel:", "", "", ""), _
        Array(4, "Art.-Nr.:", "", "EAN:", ""), _
        Array(5, "VK-Preis:", "", "EK-Preis:", ""), _
        Array(6, "Lagerort:", "", "Warengruppe:", ""), _
        Array(7, "Attribut:", "", "", "") _
    )

    Dim i As Integer
    For i = 0 To UBound(felder2)
        Dim r As Long: r = felder2(i)(0)
        wsIE.Cells(r, 2).Value = felder2(i)(1)
        wsIE.Cells(r, 2).Font.Bold = True
        wsIE.Cells(r, 2).Font.Size = 11
        wsIE.Cells(r, 3).Interior.Color = hellgrau
        wsIE.Cells(r, 3).Font.Size = 11
        wsIE.Range(wsIE.Cells(r, 3), wsIE.Cells(r, 4)).Merge
        If felder2(i)(3) <> "" Then
            wsIE.Cells(r, 5).Value = felder2(i)(3)
            wsIE.Cells(r, 5).Font.Bold = True
            wsIE.Cells(r, 5).Font.Size = 11
            wsIE.Cells(r, 6).Interior.Color = hellgrau
            wsIE.Cells(r, 6).Font.Size = 11
        End If
        wsIE.Rows(r).RowHeight = 22
    Next i

    ' Zeile 3: Artikel-Name breiter
    wsIE.Range("C3:F3").Merge

    ' Zeile 8: Leerzeile
    wsIE.Rows(8).RowHeight = 8

    ' Zeile 9: SOLL / GEZAEHLT
    wsIE.Cells(9, 2).Value = "SOLL:"
    wsIE.Cells(9, 2).Font.Bold = True: wsIE.Cells(9, 2).Font.Size = 12
    wsIE.Cells(9, 3).Interior.Color = hellgrau
    wsIE.Cells(9, 3).Font.Size = 12: wsIE.Cells(9, 3).Font.Bold = True
    wsIE.Range("C9:D9").Merge

    wsIE.Cells(9, 5).Value = "GEZAEHLT:"
    wsIE.Cells(9, 5).Font.Bold = True: wsIE.Cells(9, 5).Font.Size = 12
    wsIE.Cells(9, 6).Interior.Color = RGB(255, 255, 0)
    wsIE.Cells(9, 6).Font.Size = 14: wsIE.Cells(9, 6).Font.Bold = True
    wsIE.Rows(9).RowHeight = 26

    ' Zeile 10: Differenz-Anzeige
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
    wsIE.Range("B11:C11").Merge
    With wsIE.Cells(11, 5)
        .Value = "UEBERNEHMEN"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = gruen: .HorizontalAlignment = xlCenter
    End With
    wsIE.Range("E11:F11").Merge
    wsIE.Rows(11).RowHeight = 28

    ' Zeile 14: versteckt - Artikelzeile
    wsIE.Rows(14).Hidden = True
    wsIE.Cells(IE_ZEILE_REF, 1).Value = 0

    ' Spaltenbreiten
    wsIE.Columns(1).ColumnWidth = 2
    wsIE.Columns(2).ColumnWidth = 14
    wsIE.Columns(3).ColumnWidth = 20
    wsIE.Columns(4).ColumnWidth = 10
    wsIE.Columns(5).ColumnWidth = 14
    wsIE.Columns(6).ColumnWidth = 20

    ' Sheet-Events
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsIE.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim c As String: c = ""
    c = c & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    c = c & "    If Target.Address = ""$F$9"" Then NeueModule.InvEingabe_DiffAnzeigen" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    Cancel = True" & Chr(10)
    c = c & "    If Target.Row = 11 And Target.Column >= 2 And Target.Column <= 3 Then NeueModule.InvEingabe_Abbrechen" & Chr(10)
    c = c & "    If Target.Row = 11 And Target.Column >= 5 And Target.Column <= 6 Then NeueModule.InvEingabe_Uebernehmen" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c

    Application.ScreenUpdating = True
End Sub

' ================================================================
'  InvEingabe - Artikel befuellen (aus InvSuche aufgerufen)
' ================================================================
Sub InvEingabe_Befuellen(zeile As Long)
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Dim wsIE As Worksheet: Set wsIE = LagerMakros.GetSheet("InvEingabe")
    If wsA Is Nothing Or wsIE Is Nothing Then
        MsgBox "InvEingabe-Blatt fehlt. Bitte Setup_InvEingabe ausfuehren.", vbExclamation
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

    Dim vEAN   As String
    If colEAN > 0 Then
        Dim rawEAN As Variant: rawEAN = wsA.Cells(zeile, colEAN).Value
        If IsNumeric(rawEAN) Then
            vEAN = Format(CDbl(rawEAN), "0000000000000")
        Else
            vEAN = CStr(rawEAN)
        End If
    End If
    Dim vAnz   As Double: vAnz = IIf(colAnz > 0, Val(wsA.Cells(zeile, colAnz).Value), 0)
    Dim vEinh  As String: vEinh = IIf(colEinh > 0, wsA.Cells(zeile, colEinh).Value, "Stk")

    ' Pruefen ob schon in InvDaten vorhanden
    Dim vorhandenerIst As Double: vorhandenerIst = InvDaten_IstLaden(vEAN)

    wsIE.Cells(3, 3).Value = IIf(colArt > 0, wsA.Cells(zeile, colArt).Value, "")
    wsIE.Cells(4, 3).Value = IIf(colNr > 0, CStr(wsA.Cells(zeile, colNr).Value), "")
    wsIE.Cells(4, 6).Value = vEAN
    wsIE.Cells(5, 3).Value = IIf(colVK > 0, Format(wsA.Cells(zeile, colVK).Value, "0.00") & " EUR", "")
    wsIE.Cells(5, 6).Value = IIf(colEK > 0, Format(wsA.Cells(zeile, colEK).Value, "0.00") & " EUR", "")
    wsIE.Cells(6, 3).Value = IIf(colLag > 0, wsA.Cells(zeile, colLag).Value, "")
    wsIE.Cells(6, 6).Value = IIf(colWG > 0, wsA.Cells(zeile, colWG).Value, "")
    wsIE.Cells(7, 3).Value = IIf(colAttr > 0, wsA.Cells(zeile, colAttr).Value, "")
    wsIE.Cells(9, 3).Value = Format(vAnz, "0") & " " & vEinh
    wsIE.Cells(9, 6).Value = IIf(vorhandenerIst >= 0, vorhandenerIst, "")
    wsIE.Cells(10, 6).Value = ""
    wsIE.Cells(IE_ZEILE_REF, 1).Value = zeile

    wsIE.Visible = xlSheetVisible
    wsIE.Activate
    wsIE.Cells(9, 6).Select
End Sub

' ================================================================
'  InvEingabe - Differenz live anzeigen
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

    Dim ist  As Double: ist = Val(istStr)
    Dim soll As Double: soll = Val(sollStr)
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
    If Not wsIS Is Nothing Then
        wsIS.Activate
        wsIS.Cells(2, 2).Select
    End If
    If Not wsIE Is Nothing Then wsIE.Visible = xlSheetHidden
End Sub

' ================================================================
'  InvEingabe - Uebernehmen -> in InvDaten speichern
' ================================================================
Sub InvEingabe_Uebernehmen()
    Dim wsIE As Worksheet: Set wsIE = LagerMakros.GetSheet("InvEingabe")
    Dim wsA  As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsIE Is Nothing Then Exit Sub

    Dim istStr As String: istStr = Trim(CStr(wsIE.Cells(9, 6).Value))
    If istStr = "" Then
        MsgBox "Bitte gezaehlte Menge eingeben.", vbExclamation
        Exit Sub
    End If

    Dim zeile   As Long:   zeile = Val(wsIE.Cells(IE_ZEILE_REF, 1).Value)
    Dim ean     As String: ean = CStr(wsIE.Cells(4, 6).Value)
    Dim artNr   As String: artNr = CStr(wsIE.Cells(4, 3).Value)
    Dim artName As String: artName = wsIE.Cells(3, 3).Value
    Dim sollStr As String: sollStr = wsIE.Cells(9, 3).Value
    Dim lagerort As String: lagerort = wsIE.Cells(6, 3).Value
    Dim soll    As Double: soll = Val(sollStr)
    Dim ist     As Double: ist = Val(istStr)

    ' In InvDaten speichern
    InvDaten_Speichern ean, artNr, artName, soll, ist, lagerort

    ' Zeile in InvSuche faerben
    Dim wsIS As Worksheet: Set wsIS = LagerMakros.GetSheet("InvSuche")
    If Not wsIS Is Nothing Then
        Dim diff As Double: diff = ist - soll
        Dim lastIS As Long: lastIS = wsIS.Cells(wsIS.Rows.Count, 3).End(xlUp).Row
        Dim i As Long
        For i = 4 To lastIS
            If Val(wsIS.Cells(i, 8).Value) = zeile Then
                If diff = 0 Then
                    wsIS.Range("A" & i & ":G" & i).Interior.Color = RGB(198, 239, 206)
                ElseIf diff > 0 Then
                    wsIS.Range("A" & i & ":G" & i).Interior.Color = RGB(255, 235, 156)
                Else
                    wsIS.Range("A" & i & ":G" & i).Interior.Color = RGB(255, 199, 206)
                End If
                wsIS.Cells(i, 4).Value = ist  ' SOLL-Spalte mit IST aktualisieren
                Exit For
            End If
        Next i
    End If

    MsgBox artName & Chr(10) & "Gezaehlt: " & Format(ist, "0") & " Stk -> gespeichert", vbInformation

    ' Zurueck zu InvSuche
    InvEingabe_Abbrechen
End Sub




' ================================================================
'  ARTIKEL-SHEET TOOLBAR EINRICHTEN (Zeilen 1-4 nach Vorlage)
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
    ' Alle Datenzeilen sichtbar schalten (sauberer Ausgangszustand)
    wsA.Rows("5:5000").Hidden = False

    ' Farben gemaess Vorlage
    Dim blauDunkel As Long: blauDunkel = RGB(32, 55, 100)
    Dim gruen As Long:      gruen = RGB(55, 94, 50)
    Dim orange As Long:     orange = RGB(180, 90, 0)
    Dim grau As Long:       grau = RGB(89, 89, 89)
    Dim blauMittel As Long: blauMittel = RGB(46, 80, 144)
    Dim gitHub As Long:     gitHub = RGB(36, 41, 46)
    Dim stahlBlau As Long:  stahlBlau = RGB(31, 97, 141)
    Dim kopfgrau As Long:   kopfgrau = RGB(64, 64, 64)

    ' ==========================================
    ' ZEILE 1: Titel
    ' ==========================================
    wsA.Rows(1).RowHeight = 28
    On Error Resume Next: wsA.Range("A1:U1").Merge: On Error GoTo 0
    wsA.Cells(1, 1).Value = "Artikel - Lagerverwaltung"
    wsA.Cells(1, 1).Interior.Color = blauDunkel
    wsA.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsA.Cells(1, 1).Font.Bold = True
    wsA.Cells(1, 1).Font.Size = 14
    wsA.Cells(1, 1).HorizontalAlignment = xlCenter

    ' ==========================================
    ' ZEILE 2: Suchfeld + Shape-Buttons SUCHEN / LEEREN
    ' ==========================================
    wsA.Rows(2).RowHeight = 30
    ' ERST alte Merges aufheben – verhindert unsichtbares gelbes Feld beim 2. Run
    On Error Resume Next: wsA.Rows(2).UnMerge: On Error GoTo 0
    wsA.Rows(2).Interior.ColorIndex = xlNone

    ' A2: "Suche:" Label
    wsA.Cells(2, 1).Value = "Suche:"
    wsA.Cells(2, 1).Font.Bold = True
    wsA.Cells(2, 1).Font.Size = 14
    wsA.Cells(2, 1).Interior.Color = RGB(242, 242, 242)
    wsA.Cells(2, 1).HorizontalAlignment = xlRight

    ' B2:D2: Gelbes Suchfeld (nach UnMerge jetzt sauber neu mergen)
    On Error Resume Next: wsA.Range("B2:D2").Merge: On Error GoTo 0
    wsA.Cells(2, 2).NumberFormat = "@"
    wsA.Cells(2, 2).Interior.Color = RGB(255, 255, 153)
    wsA.Cells(2, 2).Value = ""
    wsA.Cells(2, 2).Font.Size = 14
    wsA.Cells(2, 2).Font.Color = RGB(0, 0, 0)

    ' Alle Zeile-2 Elemente sequenziell positionieren (unabhaengig von Spaltenbreiten)
    Dim xPos2  As Single: xPos2 = wsA.Range("B2:D2").Left + wsA.Range("B2:D2").Width + 3
    Dim yPos2  As Single: yPos2 = wsA.Rows(2).Top + 1
    Dim btnH2  As Single: btnH2 = wsA.Rows(2).Height - 2
    Dim btnW2  As Single: btnW2 = 75     ' SUCHEN + LEEREN
    Dim wTreffer As Single: wTreffer = 90 ' Treffer-Anzeige

    ' SUCHEN
    Dim oSuch As Shape
    Set oSuch = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos2, yPos2, btnW2, btnH2)
    With oSuch
        .Name = "btnSuchen"
        .Fill.ForeColor.RGB = gruen
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "SUCHEN"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 13
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "LagerMakros.Artikel_Suchen"
    End With
    xPos2 = xPos2 + btnW2 + 4

    ' Treffer-Anzeige als Shape (kein Zellbezug – kein Merge-Problem)
    Dim oTreffer As Shape
    Set oTreffer = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos2, yPos2, wTreffer, btnH2)
    With oTreffer
        .Name = "trefferAnzeige"
        .Fill.ForeColor.RGB = RGB(242, 242, 242)
        .Line.ForeColor.RGB = RGB(180, 180, 180)
        .Line.Visible = msoTrue
        .TextFrame.Characters.Text = "Treffer"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Color = RGB(89, 89, 89)
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
    End With
    xPos2 = xPos2 + wTreffer + 4

    ' LEEREN
    Dim oLeer As Shape
    Set oLeer = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos2, yPos2, btnW2, btnH2)
    With oLeer
        .Name = "btnLeeren"
        .Fill.ForeColor.RGB = orange
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "LEEREN"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 13
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "LagerMakros.Artikel_Suche_Leeren"
    End With
    xPos2 = xPos2 + btnW2 + 4

    ' AKTUALISIEREN
    Dim oAkt As Shape
    Set oAkt = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xPos2, yPos2, btnW2 + 30, btnH2)
    With oAkt
        .Name = "btnAktualisieren"
        .Fill.ForeColor.RGB = RGB(0, 112, 96)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "AKTUALISIEREN"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "LagerMakros.Artikel_Aktualisieren"
    End With

    ' ==========================================
    ' ZEILE 3: Toolbar als Shape-Buttons (abgerundet)
    ' Reihenfolge: GITHUB, NEUER ARTIKEL, ZU-/ABGANG, ETIKETT,
    '              EK ausbl., FILTER LOESCHEN, SCHNELLANSICHT
    ' ==========================================
    wsA.Rows(3).RowHeight = 28

    Dim btnH3 As Single: btnH3 = wsA.Rows(3).Height - 4
    Dim btnW3 As Single: btnW3 = 108   ' feste Breite pro Button
    Dim gap3  As Single: gap3 = 5      ' Abstand zwischen Buttons
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
            .Name = "btnT3_" & bi
            .Fill.ForeColor.RGB = btns(bi)(1)
            .Line.Visible = msoFalse
            .TextFrame.Characters.Text = btns(bi)(0)
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.HorizontalAlignment = xlCenter
            .TextFrame.VerticalAlignment = xlCenter
            .OnAction = btns(bi)(2)
        End With
        xPos = xPos + btnW3 + gap3
    Next bi

    ' ==========================================
    ' ZEILE 4: Spaltenkoepfe formatieren
    ' ==========================================
    wsA.Rows(4).RowHeight = 22
    With wsA.Range("A4:U4")
        .Interior.Color = kopfgrau
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
    End With

    ' ==========================================
    ' ZEILEN 1-4 EINFRIEREN
    ' ==========================================
    wsA.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    wsA.Cells(5, 1).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' ==========================================
    ' EVENT-CODE EINBETTEN
    ' ==========================================
    Dim vbA As Object: Set vbA = ThisWorkbook.VBProject.VBComponents(wsA.CodeName)
    Dim cmA As Object: Set cmA = vbA.CodeModule
    If cmA.CountOfLines > 0 Then cmA.DeleteLines 1, cmA.CountOfLines
    Dim ev As String: ev = ""
    ' Doppelklick ab Zeile 5 = ArtikelDetail
    ev = ev & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    ev = ev & "    If Target.Row >= 5 Then" & Chr(10)
    ev = ev & "        Cancel = True" & Chr(10)
    ev = ev & "        NeueModule.ArtikelDetail_Laden Target.Row" & Chr(10)
    ev = ev & "    End If" & Chr(10)
    ev = ev & "End Sub" & Chr(10)
    ev = ev & Chr(10)
    ' Aktivierung: Cursor ins Suchfeld
    ev = ev & "Private Sub Worksheet_Activate()" & Chr(10)
    ev = ev & "    On Error Resume Next" & Chr(10)
    ev = ev & "    Me.Cells(2, 2).Select" & Chr(10)
    ev = ev & "End Sub" & Chr(10)
    ev = ev & Chr(10)
    ' Suchfeld-Aenderung: Suche ausfuehren (auch per Enter)
    ev = ev & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    ev = ev & "    If Not Intersect(Target, Me.Range(""B2"")) Is Nothing Then" & Chr(10)
    ev = ev & "        On Error GoTo Fertig" & Chr(10)
    ev = ev & "        Application.EnableEvents = False" & Chr(10)
    ev = ev & "        LagerMakros.Artikel_Suchen" & Chr(10)
    ev = ev & "        Application.EnableEvents = True" & Chr(10)
    ev = ev & "    End If" & Chr(10)
    ev = ev & "    Exit Sub" & Chr(10)
    ev = ev & "Fertig: Application.EnableEvents = True" & Chr(10)
    ev = ev & "End Sub" & Chr(10)
    ev = ev & Chr(10)
    ' Klick auf Artikelzeile markieren (ab Zeile 5)
    ev = ev & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10)
    ev = ev & "    If Target.Row >= 5 Then LagerMakros.Artikel_Zeile_Markieren Target" & Chr(10)
    ev = ev & "End Sub" & Chr(10)
    cmA.AddFromString ev

    Application.ScreenUpdating = True
    Application.StatusBar = "Artikel Toolbar eingerichtet."
End Sub

' ================================================================
'  ALLE SETUP-SUBS auf einmal ausfuehren
' ================================================================
Sub Setup_NeueModule()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim setupLog As String: setupLog = ""
    On Error GoTo SetupFehler

    setupLog = "ArtikelDetail"
    Setup_ArtikelDetail
    setupLog = "BewPopup"
    Setup_BewPopup
    setupLog = "InvDaten"
    Setup_InvDaten
    setupLog = "InvEingabe"
    Setup_InvEingabe
    setupLog = "Schnellansicht"
    Setup_Schnellansicht
    setupLog = "SchnellDetail"
    Setup_SchnellDetail
    setupLog = "Artikel_Toolbar"
    Setup_Artikel_Toolbar

    ' ArtikelDetail-Aufruf in Artikel-Sheet einrichten (BeforeDoubleClick)
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If Not wsA Is Nothing Then
        Dim vbA As Object: Set vbA = ThisWorkbook.VBProject.VBComponents(wsA.CodeName)
        Dim cmA As Object: Set cmA = vbA.CodeModule
        Dim vorh As Boolean: vorh = False
        Dim li As Long
        For li = 1 To cmA.CountOfLines
            If InStr(cmA.Lines(li, 1), "BeforeDoubleClick") > 0 Then vorh = True: Exit For
        Next li
        If Not vorh Then
            Dim a As String: a = ""
            a = a & Chr(10)
            a = a & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
            a = a & "    If Target.Row >= 3 Then" & Chr(10)
            a = a & "        Cancel = True" & Chr(10)
            a = a & "        NeueModule.ArtikelDetail_Laden Target.Row" & Chr(10)
            a = a & "    End If" & Chr(10)
            a = a & "End Sub" & Chr(10)
            cmA.AddFromString a
        End If
    End If

    ' InvSuche: ArtikelWaehlen auf InvEingabe umleiten
    Dim wsIS As Worksheet: Set wsIS = LagerMakros.GetSheet("InvSuche")
    If Not wsIS Is Nothing Then
        Dim vbIS As Object: Set vbIS = ThisWorkbook.VBProject.VBComponents(wsIS.CodeName)
        Dim cmIS As Object: Set cmIS = vbIS.CodeModule
        Dim li2 As Long, gefunden As Boolean: gefunden = False
        For li2 = 1 To cmIS.CountOfLines
            If InStr(cmIS.Lines(li2, 1), "InvEingabe_Befuellen") > 0 Then gefunden = True: Exit For
        Next li2
        If Not gefunden Then
            Dim ii As Long
            For ii = 1 To cmIS.CountOfLines
                If InStr(cmIS.Lines(ii, 1), "InvSuche_ArtikelWaehlen") > 0 Then
                    cmIS.ReplaceLine ii, Replace(cmIS.Lines(ii, 1), _
                        "LagerMakros.InvSuche_ArtikelWaehlen Target.Row", _
                        "NeueModule.InvEingabe_Befuellen Val(Me.Cells(Target.Row, 8).Value)")
                    Exit For
                End If
            Next ii
        End If
    End If

    MsgBox "Setup abgeschlossen!" & Chr(10) & Chr(10) & _
           "Neue Funktionen:" & Chr(10) & _
           "- Doppelklick auf Artikel -> Detailansicht" & Chr(10) & _
           "- InvSuche -> Artikel anklicken -> Inventureingabe" & Chr(10) & _
           "- InvDaten: Inventur sitzungsuebergreifend speichern" & Chr(10) & _
           "- Inventur abschliessen: Makro 'InvDaten_Abschliessen'", _
           vbInformation, "Neue Module aktiv"
    GoTo SetupFertig
SetupFehler:
    MsgBox "Fehler in: Setup_" & setupLog & Chr(10) & Chr(10) & "Fehler " & Err.Number & ": " & Err.Description, vbCritical, "Setup-Fehler"
SetupFertig:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  EINMALIG: VK2-Spalte ins Artikel-Sheet einfuegen
' ================================================================
Sub VK2_Spalte_Einrichten()
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    If wsA Is Nothing Then MsgBox "Artikel-Sheet nicht gefunden": Exit Sub

    ' Pruefen ob VK2 schon vorhanden
    If LagerMakros.Spalte_Finden(wsA, "VK2") > 0 Then
        MsgBox "VK2-Spalte bereits vorhanden.", vbInformation
        Exit Sub
    End If

    ' VK-PREIS Spalte finden und VK2 danach einfuegen
    Dim colVK As Long: colVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    If colVK = 0 Then MsgBox "VK-PREIS Spalte nicht gefunden": Exit Sub

    wsA.Columns(colVK + 1).Insert Shift:=xlShiftToRight
    wsA.Cells(2, colVK + 1).Value = "VK2"
    wsA.Cells(2, colVK + 1).Font.Bold = True
    wsA.Cells(2, colVK + 1).Font.Size = 11
    wsA.Columns(colVK + 1).ColumnWidth = 10

    ' Gleiche Hintergrundfarbe wie Nachbarzelle
    On Error Resume Next
    wsA.Cells(2, colVK + 1).Interior.Color = wsA.Cells(2, colVK).Interior.Color
    On Error GoTo 0

    MsgBox "VK2-Spalte eingefuegt nach Spalte " & colVK & " (VK-PREIS).", vbInformation
End Sub

' ================================================================
'  SCHNELLANSICHT - OEFFNEN / SCHLIESSEN / EK-TOGGLE
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
    wsS.Cells(4, 1).Select          ' nach oben scrollen
    Application.ScreenUpdating = True
    Application.StatusBar = "Schnellansicht aktualisiert."
End Sub

Sub Schnellansicht_Schliessen()
    Dim wsA As Worksheet: Set wsA = LagerMakros.GetSheet("Artikel")
    Application.StatusBar = False
    If Not wsA Is Nothing Then wsA.Activate
End Sub

Sub Schnellansicht_EK_Toggle()
    ' EK-Preis ist jetzt in Spalte G (7)
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = wsS.Cells(wsS.Rows.Count, 4).End(xlUp).Row  ' Artikel-Spalte D
    If lastRow < 4 Then lastRow = 4

    Dim rngEK As Range
    Set rngEK = wsS.Range(wsS.Cells(4, 7), wsS.Cells(lastRow, 7))  ' Spalte G = EK

    If rngEK.Cells(1, 1).NumberFormat = ";;;" Then
        ' Einblenden
        rngEK.NumberFormat = "#,##0.00 " & ChrW(8364)
        wsS.Cells(3, 7).Value = "EK-Preis " & ChrW(8364)
        wsS.Cells(2, 7).Value = "EK AUSBL."
    Else
        ' Ausblenden
        rngEK.NumberFormat = ";;;"
        wsS.Cells(3, 7).Value = ""
        wsS.Cells(2, 7).Value = "EK EINBL."
    End If
End Sub

' ================================================================
'  SCHNELLANSICHT - SETUP (einmalig, Formatierung + Events)
' ================================================================

Sub Setup_Schnellansicht()
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If wsS Is Nothing Then
        MsgBox "Schnellansicht-Sheet nicht gefunden!", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    wsS.Visible = xlSheetVisible

    ' Bestehende Zusammenfuehrungen aufheben
    On Error Resume Next
    wsS.Cells.UnMerge
    On Error GoTo 0

    ' --- Zeile 1: Titelzeile ---
    wsS.Rows(1).RowHeight = 22
    On Error Resume Next: wsS.Range("A1:L1").Merge: On Error GoTo 0
    wsS.Cells(1, 1).Value = "Schnellansicht - Lagerverwaltung"
    wsS.Cells(1, 1).Font.Bold = True
    wsS.Cells(1, 1).Font.Size = 13
    wsS.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsS.Cells(1, 1).Interior.Color = RGB(31, 73, 125)
    wsS.Cells(1, 1).HorizontalAlignment = xlLeft

    ' --- Zeile 2: Toolbar ---
    wsS.Rows(2).RowHeight = 22
    ' A2: Suche-Label
    wsS.Cells(2, 1).Value = "Suche:"
    wsS.Cells(2, 1).Font.Bold = True
    wsS.Cells(2, 1).Interior.ColorIndex = xlNone
    ' B2:C2 merged: Suchfeld (gelb, als TEXT formatieren → verhindert Datum-Autokonvert.)
    On Error Resume Next: wsS.Range("B2:C2").Merge: On Error GoTo 0
    wsS.Cells(2, 2).NumberFormat = "@"
    wsS.Cells(2, 2).Interior.Color = RGB(255, 255, 153)
    wsS.Cells(2, 2).Value = ""
    ' D2: SUCHEN (gruen)
    wsS.Cells(2, 4).Value = "SUCHEN"
    wsS.Cells(2, 4).Font.Bold = True
    wsS.Cells(2, 4).Interior.Color = RGB(0, 176, 80)
    wsS.Cells(2, 4).Font.Color = RGB(255, 255, 255)
    wsS.Cells(2, 4).HorizontalAlignment = xlCenter
    ' E2: FILTER LOESCHEN (orange)
    wsS.Cells(2, 5).Value = "FILTER LOESCHEN"
    wsS.Cells(2, 5).Font.Bold = True
    wsS.Cells(2, 5).Interior.Color = RGB(237, 125, 49)
    wsS.Cells(2, 5).Font.Color = RGB(255, 255, 255)
    wsS.Cells(2, 5).HorizontalAlignment = xlCenter
    ' F2: AKTUALISIEREN (blau)
    wsS.Cells(2, 6).Value = "AKTUALISIEREN"
    wsS.Cells(2, 6).Font.Bold = True
    wsS.Cells(2, 6).Interior.Color = RGB(68, 114, 196)
    wsS.Cells(2, 6).Font.Color = RGB(255, 255, 255)
    wsS.Cells(2, 6).HorizontalAlignment = xlCenter
    ' G2: EK AUSBL. (blau-grau)
    wsS.Cells(2, 7).Value = "EK AUSBL."
    wsS.Cells(2, 7).Font.Bold = True
    wsS.Cells(2, 7).Interior.Color = RGB(91, 155, 213)
    wsS.Cells(2, 7).Font.Color = RGB(255, 255, 255)
    wsS.Cells(2, 7).HorizontalAlignment = xlCenter
    ' H2:I2 merged: Treffer-Anzeige (hellgrau)
    On Error Resume Next: wsS.Range("H2:I2").Merge: On Error GoTo 0
    wsS.Cells(2, 8).Value = ""
    wsS.Cells(2, 8).Interior.Color = RGB(217, 217, 217)
    wsS.Cells(2, 8).Font.Bold = True
    wsS.Cells(2, 8).HorizontalAlignment = xlCenter
    ' J2: SCHLIESSEN (rot)
    wsS.Cells(2, 10).Value = "SCHLIESSEN"
    wsS.Cells(2, 10).Font.Bold = True
    wsS.Cells(2, 10).Interior.Color = RGB(192, 0, 0)
    wsS.Cells(2, 10).Font.Color = RGB(255, 255, 255)
    wsS.Cells(2, 10).HorizontalAlignment = xlCenter
    ' K2, L2: leer
    wsS.Cells(2, 11).Interior.ColorIndex = xlNone: wsS.Cells(2, 11).Value = ""
    wsS.Cells(2, 12).Interior.ColorIndex = xlNone: wsS.Cells(2, 12).Value = ""

    ' --- Zeile 3: Spaltenkoepfe ---
    ' Vorlage: A=spacer, B=#, C=Art.-Nr., D=Artikel, E=EAN,
    '          F=VK-Preis, G=EK-Preis, H=Bestand, I=Einheit,
    '          J=Lagerort, K=Warengruppe, L=Attribut, M=Zeilenverweis(hidden)
    wsS.Rows(3).RowHeight = 18
    wsS.Cells(3, 1).Value = ""
    wsS.Cells(3, 2).Value = "#"
    wsS.Cells(3, 3).Value = "Art.-Nr."
    wsS.Cells(3, 4).Value = "Artikel"
    wsS.Cells(3, 5).Value = "EAN"
    wsS.Cells(3, 6).Value = "VK-Preis " & ChrW(8364)
    wsS.Cells(3, 7).Value = "EK-Preis " & ChrW(8364)
    wsS.Cells(3, 8).Value = "Bestand"
    wsS.Cells(3, 9).Value = "Einheit"
    wsS.Cells(3, 10).Value = "Lagerort"
    wsS.Cells(3, 11).Value = "Warengruppe"
    wsS.Cells(3, 12).Value = "Attribut"
    wsS.Cells(3, 13).Value = ""
    With wsS.Range("A3:L3")
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
        .HorizontalAlignment = xlCenter
    End With

    ' --- Spaltenbreiten ---
    wsS.Columns(1).ColumnWidth = 2    ' Spacer
    wsS.Columns(2).ColumnWidth = 5    ' #
    wsS.Columns(3).ColumnWidth = 13   ' Art.-Nr.
    wsS.Columns(4).ColumnWidth = 32   ' Artikel
    wsS.Columns(5).ColumnWidth = 16   ' EAN
    wsS.Columns(6).ColumnWidth = 11   ' VK-Preis
    wsS.Columns(7).ColumnWidth = 11   ' EK-Preis (ausblendbar)
    wsS.Columns(8).ColumnWidth = 8    ' Bestand
    wsS.Columns(9).ColumnWidth = 8    ' Einheit
    wsS.Columns(10).ColumnWidth = 13  ' Lagerort
    wsS.Columns(11).ColumnWidth = 14  ' Warengruppe
    wsS.Columns(12).ColumnWidth = 12  ' Attribut
    wsS.Columns(13).Hidden = True     ' Zeilenverweis (hidden)

    ' --- Art.-Nr. und EAN als Text formatieren (keine wissenschaftliche Notation) ---
    wsS.Columns(3).NumberFormat = "@"  ' Art.-Nr.
    wsS.Columns(5).NumberFormat = "@"  ' EAN

    ' --- Fenster fixieren (Zeilen 1-3 einfrieren) ---
    wsS.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    wsS.Cells(4, 1).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' --- Event-Code in Sheet einbetten ---
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
    c = c & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10)
    c = c & "    LagerMakros.Schnellansicht_Handler Target" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "" & Chr(10)
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    If Target.Row >= 4 And Target.Column >= 2 And Target.Column <= 12 Then" & Chr(10)
    c = c & "        Cancel = True" & Chr(10)
    c = c & "        Dim sEAN As String: sEAN = CStr(Me.Cells(Target.Row, 5).Value)" & Chr(10)
    c = c & "        Dim sArt As String: sArt = CStr(Me.Cells(Target.Row, 4).Value)" & Chr(10)
    c = c & "        If sArt <> """" Then NeueModule.SchnellDetail_Laden sEAN, sArt" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cmS.AddFromString c

    Application.ScreenUpdating = True
    Application.StatusBar = "Setup Schnellansicht abgeschlossen."
End Sub

' ================================================================
'  SCHNELLDETAIL - Einfache Detailansicht (aus Schnellansicht)
'  Zeigt: EAN, Artikel, VK-Preis, Menge, EK-Preis
' ================================================================

Sub SchnellDetail_Laden(sEAN As String, sArtName As String)
    ' Sheet suchen
    Dim wsSD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "SchnellDetail" Then Set wsSD = ws: Exit For
    Next ws
    If wsSD Is Nothing Then
        MsgBox "SchnellDetail-Sheet fehlt. Bitte Setup_NeueModule ausfuehren.", vbExclamation
        Exit Sub
    End If

    ' Artikel im Artikel-Sheet suchen
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
    Dim colAttr As Long: colAttr = LagerMakros.Spalte_Finden(wsA, "ATTRIBUT")
    If colArt = 0 Then GoTo SchnellDetailOeffnen  ' Sheet trotzdem oeffnen

    Dim lastRow As Long: lastRow = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
    Dim gefunden As Long: gefunden = 0
    Dim i As Long

    ' Erst per EAN suchen (genauer), dann per Name
    If sEAN <> "" And colEAN > 0 Then
        For i = 5 To lastRow
            If CStr(wsA.Cells(i, colEAN).Value) = sEAN Then gefunden = i: Exit For
        Next i
    End If
    If gefunden = 0 And sArtName <> "" Then
        For i = 5 To lastRow
            If CStr(wsA.Cells(i, colArt).Value) = sArtName Then gefunden = i: Exit For
        Next i
    End If

    If gefunden > 0 Then
        ' Werte laden und anzeigen
        Dim vEAN  As String: If colEAN > 0 Then vEAN = CStr(wsA.Cells(gefunden, colEAN).Value)
        Dim vNr   As String: If colNr > 0 Then vNr = CStr(wsA.Cells(gefunden, colNr).Value)
        Dim vVK   As Double: If colVK > 0 Then vVK = Val(wsA.Cells(gefunden, colVK).Value)
        Dim vEK   As Double: If colEK > 0 Then vEK = Val(wsA.Cells(gefunden, colEK).Value)
        Dim vAnz  As Double: If colAnz > 0 Then vAnz = Val(wsA.Cells(gefunden, colAnz).Value)
        Dim vEinh As String: If colEinh > 0 Then vEinh = CStr(wsA.Cells(gefunden, colEinh).Value)
        Dim vLag  As String: If colLag > 0 Then vLag = CStr(wsA.Cells(gefunden, colLag).Value)
        Dim vWG   As String: If colWG > 0 Then vWG = CStr(wsA.Cells(gefunden, colWG).Value)
        Dim vAttr As String: If colAttr > 0 Then vAttr = CStr(wsA.Cells(gefunden, colAttr).Value)

        ' Anzeige befuellen (Zeilen 3-12, Wertfeld = Spalte C)
        wsSD.Cells(3, 3).Value = sArtName                    ' Artikel
        wsSD.Cells(4, 3).Value = vNr                         ' Art.-Nr.
        wsSD.Cells(5, 3).NumberFormat = "@"
        wsSD.Cells(5, 3).Value = vEAN                        ' EAN (Text)
        wsSD.Cells(6, 3).Value = vVK                         ' VK-Preis
        wsSD.Cells(6, 3).NumberFormat = "#,##0.00"
        wsSD.Cells(7, 3).Value = vEK                         ' EK-Preis
        wsSD.Cells(7, 3).NumberFormat = "#,##0.00"
        wsSD.Cells(8, 3).Value = vAnz                        ' Bestand
        wsSD.Cells(8, 3).NumberFormat = "0"
        wsSD.Cells(9, 3).Value = vEinh                       ' Einheit
        wsSD.Cells(10, 3).Value = vLag                       ' Lagerort
        wsSD.Cells(11, 3).Value = vWG                        ' Warengruppe
        wsSD.Cells(12, 3).Value = vAttr                      ' Attribut
        ' EK-Zeile standardmaessig ausblenden
        wsSD.Rows(7).Hidden = True
    Else
        ' Artikel nicht gefunden – trotzdem Sheet oeffnen mit Artikelname
        wsSD.Cells(3, 3).Value = sArtName
        wsSD.Cells(4, 3).Value = ""
        wsSD.Cells(5, 3).Value = sEAN
        wsSD.Range("D6:D12").Value = "?"
    End If

SchnellDetailOeffnen:
    ' EK-Button-Text auf "EK EINBL." setzen (EK ist standardmaessig versteckt)
    Dim shpSD As Shape
    For Each shpSD In wsSD.Shapes
        If shpSD.Name = "btnEKToggle" Then
            shpSD.TextFrame.Characters.Text = "EK EINBL."
            Exit For
        End If
    Next shpSD

    ' Schnellansicht verstecken, SchnellDetail zeigen
    Dim wsS As Worksheet: Set wsS = LagerMakros.GetSheet("Schnell")
    If Not wsS Is Nothing Then wsS.Visible = xlSheetHidden
    wsSD.Visible = xlSheetVisible
    wsSD.Activate
    Exit Sub

SchnellDetailFehler:
    ' Bei Fehler trotzdem Sheet oeffnen
    MsgBox "Fehler " & Err.Number & " beim Laden: " & Err.Description, vbExclamation, "SchnellDetail"
    Resume SchnellDetailOeffnen
End Sub

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

Sub Setup_SchnellDetail()
    ' Sheet erstellen falls nicht vorhanden
    Dim wsSD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "SchnellDetail" Then Set wsSD = ws: Exit For
    Next ws
    If wsSD Is Nothing Then
        Set wsSD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSD.Name = "SchnellDetail"
    End If

    wsSD.Cells.Clear
    wsSD.Cells.Interior.ColorIndex = xlNone

    ' Zeile 1: Titel
    wsSD.Range("A1:E1").Merge
    wsSD.Cells(1, 1).Value = "Artikel - Schnellansicht"
    wsSD.Cells(1, 1).Interior.Color = RGB(31, 73, 125)
    wsSD.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    wsSD.Cells(1, 1).Font.Bold = True
    wsSD.Cells(1, 1).Font.Size = 14
    wsSD.Cells(1, 1).HorizontalAlignment = xlCenter
    wsSD.Rows(1).RowHeight = 30

    ' Zeile 2: Platz fuer SCHLIESSEN-Button (Shape wird spaeter eingefuegt)
    wsSD.Rows(2).RowHeight = 28

    ' Zeile 3-12: Felder (10 Felder wie Schnellansicht-Vorlage)
    Dim felder As Variant
    felder = Array( _
        "Artikel:", _
        "Art.-Nr.:", _
        "EAN:", _
        "VK-Preis " & ChrW(8364) & ":", _
        "EK-Preis " & ChrW(8364) & ":", _
        "Bestand:", _
        "Einheit:", _
        "Lagerort:", _
        "Warengruppe:", _
        "Attribut:")
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

    ' Spaltenbreiten anpassen (EAN-Zeile braucht mehr Platz)
    wsSD.Columns(6).ColumnWidth = 6

    ' Alte Shapes entfernen
    Dim shp As Shape
    For Each shp In wsSD.Shapes
        shp.Delete
    Next shp

    ' Button 1: EK AUSBL. (blau-grau) an B2:C2
    Dim rEK As Range: Set rEK = wsSD.Range("B2:C2")
    Dim oEK As Shape
    Set oEK = wsSD.Shapes.AddShape(msoShapeRectangle, _
        rEK.Left, rEK.Top, rEK.Width, rEK.Height)
    With oEK
        .Name = "btnEKToggle"
        .Fill.ForeColor.RGB = RGB(91, 155, 213)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = "EK AUSBL."
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .OnAction = "NeueModule.SchnellDetail_EK_Toggle"
    End With

    ' Button 2: SCHLIESSEN (rot) an D2:E2
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

    ' EAN-Feld (Zeile 5, Spalte C) als Text formatieren
    wsSD.Cells(5, 3).NumberFormat = "@"

    wsSD.Visible = xlSheetHidden
    Application.StatusBar = "Setup SchnellDetail abgeschlossen."
End Sub

' ================================================================
'  SCHNELLDETAIL - EK-PREIS EIN/AUSBLENDEN
' ================================================================
Sub SchnellDetail_EK_Toggle()
    Dim wsSD As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "SchnellDetail" Then Set wsSD = ws: Exit For
    Next ws
    If wsSD Is Nothing Then Exit Sub

    ' EK-Preis ist in Zeile 7
    Dim istVersteckt As Boolean: istVersteckt = wsSD.Rows(7).Hidden
    wsSD.Rows(7).Hidden = Not istVersteckt

    ' Button-Text anpassen
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
' ============================================================
' Diese Subs in dein bestehendes Modul "NeueModule" einfuegen
' Einfach ans Ende des Moduls kopieren
' ============================================================




' ------------------------------------------------------------
' Hauptroutine � wird vom Button "GITHUB" aufgerufen
' ------------------------------------------------------------
Sub GitHub_Export()

    On Error GoTo Fehler
    ' === KONFIGURATION � hier anpassen falls noetig ===
Const GIT_DIR     As String = "D:\_KI-Projekte-2026\07_Lagerverwaltung-Excel\Lagerverw. FB300\"
Const LAGER_SHEET As String = "Schnellansicht"
' ===================================================

    ' --- 1. VBA-Module als .bas exportieren ---
    Dim moduleNames As Variant
    moduleNames = Array("LagerMakros", "NeueModule")

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

    ' --- 2. Lagerdaten als JSON exportieren ---
    Call ExportLagerJSON(GIT_DIR, LAGER_SHEET)

    ' --- 2b. Artikel-JSON f�r PAM Mobil exportieren (kompaktes Format, UTF-8) ---
    Call ExportArtikelPAMJson(LAGER_SHEET)

    ' --- 3. Git: add -> commit -> push ---
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim ts As String
    ts = Format(Now(), "DD.MM.YYYY HH\:MM")

    Dim commitMsg As String
    commitMsg = "Lagerverwaltung Update - " & ts

    Dim gitCmd As String
    gitCmd = "cmd /c cd /d """ & GIT_DIR & """" & _
             " && git add -A" & _
             " && git commit -m """ & commitMsg & """" & _
             " && git push"

    Dim rc As Long
    rc = sh.Run(gitCmd, 1, True)

    If rc = 0 Then
        MsgBox "Erfolgreich nach GitHub gepusht!" & vbNewLine & _
               "Commit: " & commitMsg, _
               vbInformation, "GitHub Export"
    Else
        MsgBox "Git meldete Fehler-Code " & rc & "." & vbNewLine & _
               "Bitte pruefe das CMD-Fenster.", _
               vbExclamation, "GitHub Export - Warnung"
    End If

    Exit Sub

Fehler:
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "GitHub Export - Fehler"

End Sub


' ------------------------------------------------------------
' Exportiert Blatt "Schnellansicht" als lager.json
' Spalten: # | Art.-Nr. | Artikel | EAN | VK | EK |
'          Bestand | Einheit | Lagerort | Warengruppe | Attribut
' ------------------------------------------------------------
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
    lastRow = ws.Cells(ws.Rows.Count, C_ARTNR).End(xlUp).Row

    Dim json    As String
    Dim isFirst As Boolean
    isFirst = True
    json = "[" & vbNewLine

    Dim i     As Long
    Dim artNr As String

    For i = 3 To lastRow

        artNr = Trim(CStr(ws.Cells(i, C_ARTNR).Value))
        If artNr = "" Then GoTo NextRow

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
            """attribut"":" & JStr(ws.Cells(i, C_ATTRIBUT)) & _
            "}"

NextRow:
    Next i

    json = json & vbNewLine & "]"

    Dim fNum As Integer
    fNum = FreeFile
    Open targetDir & "lager.json" For Output As #fNum
    Print #fNum, json
    Close #fNum

End Sub


' ------------------------------------------------------------
' Exportiert Artikel-Liste f�r PAM Mobil als artikel.json
' Format: [["Artikelname","Einheit","Warengruppe"], ...]
' Ziel:   D:\_KI-Projekte-2026\03_PAM-Mobil\artikel.json
' Kodierung: UTF-8 via ADODB.Stream (Umlaute korrekt!)
' ------------------------------------------------------------
Private Sub ExportArtikelPAMJson(sheetName As String)

    Const PAM_DIR As String = "D:\_KI-Projekte-2026\03_PAM-Mobil\"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Const C_ARTNR       As Integer = 3
    Const C_ARTIKEL     As Integer = 4
    Const C_EINHEIT     As Integer = 9
    Const C_WARENGRUPPE As Integer = 11

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, C_ARTNR).End(xlUp).Row

    Dim json    As String
    Dim isFirst As Boolean
    isFirst = True
    json = "["

    Dim i     As Long
    Dim artNr As String
    Dim artName As String
    Dim einheit As String
    Dim gruppe  As String

    For i = 3 To lastRow
        artNr = Trim(CStr(ws.Cells(i, C_ARTNR).Value))
        If artNr = "" Then GoTo NextRowPAM

        artName = Trim(CStr(ws.Cells(i, C_ARTIKEL).Value))
        einheit = Trim(CStr(ws.Cells(i, C_EINHEIT).Value))
        gruppe  = Trim(CStr(ws.Cells(i, C_WARENGRUPPE).Value))

        If artName = "" Then GoTo NextRowPAM

        If Not isFirst Then json = json & ","
        isFirst = False

        json = json & "[" & _
            """" & JStrPAM(artName) & """," & _
            """" & JStrPAM(einheit) & """," & _
            """" & JStrPAM(gruppe) & """" & _
            "]"

NextRowPAM:
    Next i

    json = json & "]"

    ' UTF-8 speichern via ADODB.Stream
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type    = 2  ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText json
    stm.SaveToFile PAM_DIR & "artikel.json", 2  ' adSaveCreateOverWrite
    stm.Close
    Set stm = Nothing

End Sub


' Hilfsfunktion: String f�r JSON escapen (ohne umgebende Anf�hrungszeichen)
Private Function JStrPAM(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    JStrPAM = s
End Function


' Hilfsfunktion: Zellwert als JSON-String
Private Function JStr(cell As Range) As String
    Dim s As String
    s = CStr(cell.Value)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    JStr = """" & s & """"
End Function


