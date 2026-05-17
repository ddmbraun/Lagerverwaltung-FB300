Attribute VB_Name = "Modul1"
Sub SetupNeuenArtikelSheet()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Neuer Artikel").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim wsNA As Worksheet
    Set wsNA = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Suche & Scanner"))
    wsNA.Name = "Neuer Artikel"

    With wsNA.Range("A1")
        .Value = "NEUEN ARTIKEL ANLEGEN"
        .Font.Bold = True: .Font.Size = 14: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(32, 55, 100)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    wsNA.Rows(1).RowHeight = 30
    wsNA.Range("A1:C1").Merge

    wsNA.Range("A2") = "Felder ausfuellen - dann SPEICHERN doppelklicken!"
    wsNA.Range("A2").Font.Italic = True
    wsNA.Range("A2").Font.Color = RGB(192, 0, 0)
    wsNA.Range("A2:C2").Merge

    Dim r As Integer, labels(13) As String
    labels(0) = "Artikelnr.:": labels(1) = "Artikel:": labels(2) = "EK-Preis Euro:"
    labels(3) = "VK-Preis Euro:": labels(4) = "Warengruppe:": labels(5) = "Lagerort:"
    labels(6) = "Lieferant:": labels(7) = "MWST %:": labels(8) = "EAN13:"
    labels(9) = "Variablesfeld:": labels(10) = "Attribut:": labels(11) = "TextA:"
    labels(12) = "TextB:": labels(13) = "Vermerk:"

    For r = 0 To 13
        wsNA.Cells(r + 4, 1).Value = labels(r)
        wsNA.Cells(r + 4, 1).Font.Bold = True
        wsNA.Cells(r + 4, 1).Font.Size = 11
        wsNA.Cells(r + 4, 2).Interior.Color = RGB(255, 255, 0)
        wsNA.Cells(r + 4, 2).Font.Size = 11
        wsNA.Rows(r + 4).RowHeight = 22
    Next r

    With wsNA.Range("A19")
        .Value = "SPEICHERN (Doppelklick!)"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(55, 86, 35): .HorizontalAlignment = xlCenter
    End With
    wsNA.Rows(19).RowHeight = 28

    With wsNA.Range("B19")
        .Value = "FELDER LEEREN (Doppelklick!)"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(192, 0, 0): .HorizontalAlignment = xlCenter
    End With

    wsNA.Columns("A").ColumnWidth = 18
    wsNA.Columns("B").ColumnWidth = 35
    wsNA.Columns("C").ColumnWidth = 30

    MsgBox "Sheet 'Neuer Artikel' wurde erstellt!", vbInformation
End Sub

' ================================================================
'  ARTIKEL SPEICHERN - dynamische Spaltenerkennung (robust)
' ================================================================
Sub ArtikelSpeichern()
    Dim wsNA As Worksheet, wsA As Worksheet
    Set wsNA = ThisWorkbook.Sheets("Neuer Artikel")
    Set wsA = LagerMakros.GetSheet("Artikel")

    If wsA Is Nothing Then
        MsgBox "Artikel-Sheet nicht gefunden!", vbCritical
        Exit Sub
    End If

    If Trim(wsNA.Range("B4").Value) = "" Then
        MsgBox "Bitte Artikelnummer eingeben!", vbExclamation
        Exit Sub
    End If

    ' Spalten dynamisch ermitteln (robust gegen Umstrukturierungen)
    Dim colNr     As Long: colNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    Dim colArt    As Long: colArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim colVK     As Long: colVK = LagerMakros.Spalte_Finden(wsA, "VK-PREIS")
    Dim colEK     As Long: colEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim colAnz    As Long: colAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim colEinh   As Long: colEinh = LagerMakros.Spalte_Finden(wsA, "EINHEIT")
    Dim colEAN    As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colWG     As Long: colWG = LagerMakros.Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colLag    As Long: colLag = LagerMakros.Spalte_Finden(wsA, "LAGERORT")
    Dim colLief   As Long: colLief = LagerMakros.Spalte_Finden(wsA, "LIEFERANT")
    Dim colMwst   As Long: colMwst = LagerMakros.Spalte_Finden(wsA, "MWST")
    Dim colAttr   As Long: colAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")
    Dim colTextA  As Long: colTextA = LagerMakros.Spalte_Finden(wsA, "TextA")
    Dim colTextB  As Long: colTextB = LagerMakros.Spalte_Finden(wsA, "TextB")
    Dim colVerm   As Long: colVerm = LagerMakros.Spalte_Finden(wsA, "VERMERK")
    Dim colWEDat  As Long: colWEDat = LagerMakros.Spalte_Finden(wsA, "WE-Datum")

    Dim nextRow As Long
    nextRow = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row + 1

    ' Format aus Zeile 3 kopieren
    wsA.Rows(3).Copy
    wsA.Rows(nextRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Werte schreiben - nur in vorhandene Spalten
    If colNr > 0 Then wsA.Cells(nextRow, colNr).Value = CStr(wsNA.Range("B4").Value)
    If colArt > 0 Then wsA.Cells(nextRow, colArt).Value = wsNA.Range("B5").Value
    If colEK > 0 Then wsA.Cells(nextRow, colEK).Value = wsNA.Range("B6").Value
    If colVK > 0 Then wsA.Cells(nextRow, colVK).Value = wsNA.Range("B7").Value
    If colWG > 0 Then wsA.Cells(nextRow, colWG).Value = wsNA.Range("B8").Value
    If colLag > 0 Then wsA.Cells(nextRow, colLag).Value = wsNA.Range("B9").Value
    If colLief > 0 Then wsA.Cells(nextRow, colLief).Value = wsNA.Range("B10").Value
    If colMwst > 0 Then wsA.Cells(nextRow, colMwst).Value = wsNA.Range("B11").Value
    If colEAN > 0 Then
        wsA.Cells(nextRow, colEAN).NumberFormat = "@"
        wsA.Cells(nextRow, colEAN).Value = CStr(wsNA.Range("B12").Value)
    End If
    If colAttr > 0 Then wsA.Cells(nextRow, colAttr).Value = wsNA.Range("B14").Value
    If colTextA > 0 Then wsA.Cells(nextRow, colTextA).Value = wsNA.Range("B15").Value
    If colTextB > 0 Then wsA.Cells(nextRow, colTextB).Value = wsNA.Range("B16").Value
    If colVerm > 0 Then wsA.Cells(nextRow, colVerm).Value = wsNA.Range("B17").Value
    If colAnz > 0 Then wsA.Cells(nextRow, colAnz).Value = wsNA.Range("B18").Value
    If colWEDat > 0 Then wsA.Cells(nextRow, colWEDat).Value = wsNA.Range("B19").Value

    ' Filter entfernen, zur neuen Zeile springen
    On Error Resume Next
    wsA.ShowAllData
    On Error GoTo 0
    wsA.Activate
    Application.GoTo wsA.Range("A" & nextRow), True

    MsgBox "Artikel '" & wsNA.Range("B5").Value & "' gespeichert!", vbInformation
    FelderLeeren
End Sub

Sub FelderLeeren()
    Dim wsNA As Worksheet
    Set wsNA = ThisWorkbook.Sheets("Neuer Artikel")
    wsNA.Range("B4:B19").ClearContents
End Sub

Sub InventurErstellen()
    Dim wsA As Worksheet, wsI As Worksheet
    Dim lastRow As Long, nextRow As Long, i As Long
    Set wsA = LagerMakros.GetSheet("Artikel")

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Inventurliste").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsI = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsI.Name = "Inventurliste"

    wsI.Range("A1") = "INVENTURLISTE"
    wsI.Range("A1").Font.Bold = True
    wsI.Range("A1").Font.Size = 14

    Dim invDatum As String
    invDatum = InputBox("Inventurdatum eingeben:", "Inventurdatum", Format(Now, "DD.MM.YYYY"))
    wsI.Range("A2") = "Datum: " & invDatum

    wsI.Range("A4") = "ARTIKELNR": wsI.Range("B4") = "ARTIKEL"
    wsI.Range("C4") = "BESTAND": wsI.Range("D4") = "EK-PREIS": wsI.Range("E4") = "GESAMT-EK"

    Dim hCell As Range
    For Each hCell In wsI.Range("A4:E4")
        hCell.Font.Bold = True
        hCell.Interior.Color = RGB(54, 96, 146)
        hCell.Font.Color = RGB(255, 255, 255)
    Next hCell

    Dim colArt As Long: colArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim colAnz As Long: colAnz = LagerMakros.Spalte_Finden(wsA, "ANZAHL")
    Dim colEK  As Long: colEK = LagerMakros.Spalte_Finden(wsA, "EK-PREIS")
    Dim colNr  As Long: colNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")

    lastRow = wsA.Cells(wsA.Rows.Count, colArt).End(xlUp).Row
    nextRow = 5
    Dim totalEK As Double
    totalEK = 0

    For i = 3 To lastRow
        Dim bestand As Double
        bestand = Val(wsA.Cells(i, colAnz).Value)
        If bestand > 0 Then
            wsI.Cells(nextRow, 1) = wsA.Cells(i, colNr).Value
            wsI.Cells(nextRow, 2) = wsA.Cells(i, colArt).Value
            wsI.Cells(nextRow, 3) = bestand
            If colEK > 0 Then
                wsI.Cells(nextRow, 4) = wsA.Cells(i, colEK).Value
                wsI.Cells(nextRow, 5) = bestand * Val(wsA.Cells(i, colEK).Value)
                wsI.Cells(nextRow, 4).NumberFormat = "#,##0.00 EUR"
                wsI.Cells(nextRow, 5).NumberFormat = "#,##0.00 EUR"
                totalEK = totalEK + bestand * Val(wsA.Cells(i, colEK).Value)
            End If
            nextRow = nextRow + 1
        End If
    Next i

    wsI.Cells(nextRow + 1, 4) = "GESAMT EK:"
    wsI.Cells(nextRow + 1, 4).Font.Bold = True
    wsI.Cells(nextRow + 1, 5) = totalEK
    wsI.Cells(nextRow + 1, 5).NumberFormat = "#,##0.00 EUR"
    wsI.Cells(nextRow + 1, 5).Font.Bold = True

    wsI.Columns("A").ColumnWidth = 15
    wsI.Columns("B").ColumnWidth = 40
    wsI.Columns("C").ColumnWidth = 12
    wsI.Columns("D").ColumnWidth = 12
    wsI.Columns("E").ColumnWidth = 14

    MsgBox "Inventurliste mit " & (nextRow - 5) & " Artikeln erstellt!", vbInformation
    wsI.Activate
End Sub

Sub DropdownsEinrichten()
    Dim wsNA As Worksheet
    Set wsNA = ThisWorkbook.Sheets("Neuer Artikel")

    Dim wsWG As Worksheet
    Set wsWG = ThisWorkbook.Sheets("Warengruppen")
    Dim lastWG As Long
    lastWG = wsWG.Cells(wsWG.Rows.Count, "A").End(xlUp).Row
    Dim dvWG As String
    dvWG = "=Warengruppen!$A$2:$A$" & lastWG
    With wsNA.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=dvWG
        .ShowError = False
    End With
    wsNA.Range("C8") = "Dropdown aus Warengruppen"
    wsNA.Range("C8").Font.Italic = True
    wsNA.Range("C8").Font.Color = RGB(89, 89, 89)
    wsNA.Range("C8").Font.Size = 9

    Dim wsLO As Worksheet
    Set wsLO = ThisWorkbook.Sheets("Lagerorte")
    Dim lastLO As Long
    lastLO = wsLO.Cells(wsLO.Rows.Count, "A").End(xlUp).Row
    Dim dvLO As String
    dvLO = "=Lagerorte!$B$2:$B$" & lastLO
    With wsNA.Range("B9").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=dvLO
        .ShowError = False
    End With
    wsNA.Range("C9") = "Dropdown aus Lagerorte"
    wsNA.Range("C9").Font.Italic = True
    wsNA.Range("C9").Font.Color = RGB(89, 89, 89)
    wsNA.Range("C9").Font.Size = 9

    Dim wsA As Worksheet
    Set wsA = LagerMakros.GetSheet("Artikel")
    Dim lastA As Long
    lastA = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).Row

    Dim liefDict As Object
    Set liefDict = CreateObject("Scripting.Dictionary")
    Dim colLief As Long: colLief = LagerMakros.Spalte_Finden(wsA, "LIEFERANT")
    Dim j As Long
    If colLief > 0 Then
        For j = 3 To lastA
            Dim liefVal As String
            liefVal = Trim(CStr(wsA.Cells(j, colLief).Value))
            If liefVal <> "" And Not liefDict.Exists(liefVal) Then
                liefDict.Add liefVal, liefVal
            End If
        Next j
    End If

    wsNA.Range("F4:F200").ClearContents
    Dim liefList As Variant
    liefList = liefDict.Keys
    Dim k As Long
    For k = 0 To UBound(liefList)
        wsNA.Cells(k + 4, 6).Value = liefList(k)
    Next k
    With wsNA.Range("B10").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="='Neuer Artikel'!$F$4:$F$" & (UBound(liefList) + 4)
        .ShowError = False
    End With
    wsNA.Range("C10") = "Dropdown Lieferanten"
    wsNA.Range("C10").Font.Italic = True
    wsNA.Range("C10").Font.Color = RGB(89, 89, 89)
    wsNA.Range("C10").Font.Size = 9

    Dim attrDict As Object
    Set attrDict = CreateObject("Scripting.Dictionary")
    Dim colAttr As Long: colAttr = LagerMakros.Spalte_Finden(wsA, "Attribut")
    If colAttr > 0 Then
        For j = 3 To lastA
            Dim attrVal As String
            attrVal = Trim(CStr(wsA.Cells(j, colAttr).Value))
            If attrVal <> "" And Not attrDict.Exists(attrVal) Then
                attrDict.Add attrVal, attrVal
            End If
        Next j
    End If

    wsNA.Range("G4:G200").ClearContents
    Dim attrList As Variant
    attrList = attrDict.Keys
    For k = 0 To UBound(attrList)
        wsNA.Cells(k + 4, 7).Value = attrList(k)
    Next k
    With wsNA.Range("B14").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="='Neuer Artikel'!$G$4:$G$" & (UBound(attrList) + 4)
        .ShowError = False
    End With
    wsNA.Range("C14") = "Dropdown Attribute"
    wsNA.Range("C14").Font.Italic = True
    wsNA.Range("C14").Font.Color = RGB(89, 89, 89)
    wsNA.Range("C14").Font.Size = 9

    wsNA.Range("B4").NumberFormat = "@"
    wsNA.Range("B12").NumberFormat = "@"
    wsNA.Range("C4") = "Als Text eingeben (keine Umrechnung)"
    wsNA.Range("C12") = "Barcode scannen oder manuell eingeben"
    wsNA.Range("C4").Font.Italic = True: wsNA.Range("C4").Font.Size = 9
    wsNA.Range("C12").Font.Italic = True: wsNA.Range("C12").Font.Size = 9
    wsNA.Range("C12").Font.Color = RGB(0, 112, 192)

    MsgBox "Dropdowns eingerichtet! Spalten F und G sind Hilfsspalten - bitte nicht loeschen.", vbInformation
End Sub

Sub FixEANFormat()
    Dim wsA As Worksheet
    Set wsA = LagerMakros.GetSheet("Artikel")
    Dim colEAN As Long: colEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    Dim colNr  As Long: colNr = LagerMakros.Spalte_Finden(wsA, "ARTIKELNR")
    If colEAN = 0 And colNr = 0 Then MsgBox "Spalten EAN13/ARTIKELNR nicht gefunden!", vbExclamation: Exit Sub

    Dim lastRow As Long
    lastRow = wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row

    Application.ScreenUpdating = False
    Dim i As Long
    For i = 3 To lastRow
        If colEAN > 0 And wsA.Cells(i, colEAN).Value <> "" Then
            Dim ean As String
            ean = Format(CDbl(wsA.Cells(i, colEAN).Value), "0")
            wsA.Cells(i, colEAN).NumberFormat = "@"
            wsA.Cells(i, colEAN).Value = ean
        End If
        If colNr > 0 And IsNumeric(wsA.Cells(i, colNr).Value) And wsA.Cells(i, colNr).Value <> "" Then
            Dim anr As String
            anr = Format(CDbl(wsA.Cells(i, colNr).Value), "0")
            wsA.Cells(i, colNr).NumberFormat = "@"
            wsA.Cells(i, colNr).Value = anr
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "EAN und Artikelnr. wurden als Text formatiert!", vbInformation
End Sub

Sub FelderErgaenzen()
    Dim wsNA As Worksheet
    Set wsNA = ThisWorkbook.Sheets("Neuer Artikel")

    With wsNA.Range("A22")
        .Value = "SPEICHERN (Doppelklick!)"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(55, 86, 35): .HorizontalAlignment = xlCenter
    End With
    wsNA.Rows(22).RowHeight = 28
    With wsNA.Range("B22")
        .Value = "FELDER LEEREN (Doppelklick!)"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(192, 0, 0): .HorizontalAlignment = xlCenter
    End With
    wsNA.Range("A19").ClearContents: wsNA.Range("A19").Interior.ColorIndex = xlNone
    wsNA.Range("B19").ClearContents: wsNA.Range("B19").Interior.ColorIndex = xlNone

    wsNA.Range("A18") = "Menge (Bestand):"
    wsNA.Range("A18").Font.Bold = True: wsNA.Range("A18").Font.Size = 11
    wsNA.Range("B18").Interior.Color = RGB(255, 255, 0)
    wsNA.Range("B18").Font.Size = 11
    wsNA.Range("C18") = "Anfangsbestand eingeben (z.B. 10)"
    wsNA.Range("C18").Font.Italic = True: wsNA.Range("C18").Font.Size = 9
    wsNA.Rows(18).RowHeight = 22

    wsNA.Range("A19") = "Wareneingang-Datum:"
    wsNA.Range("A19").Font.Bold = True: wsNA.Range("A19").Font.Size = 11
    wsNA.Range("B19").Interior.Color = RGB(255, 255, 0)
    wsNA.Range("B19").Font.Size = 11
    wsNA.Range("B19").NumberFormat = "DD.MM.YYYY"
    wsNA.Range("C19") = "z.B. 10.04.2026"
    wsNA.Range("C19").Font.Italic = True: wsNA.Range("C19").Font.Size = 9
    wsNA.Rows(19).RowHeight = 22

    MsgBox "Felder hinzugefuegt!", vbInformation
End Sub

Sub InventurButtonErstellen()
    Dim ws As Worksheet
    Dim shp As Object
    Dim rng As Range

    Set ws = LagerMakros.GetSheet("Artikel")

    On Error Resume Next
    ws.Shapes("InventurButton").Delete
    On Error GoTo 0

    Set rng = ws.Range("E1:G1")
    Set shp = ws.Shapes.AddShape(5, rng.Left, rng.Top, rng.Width, 30)

    With shp
        .Name = "InventurButton"
        .Text = "INVENTUR ERSTELLEN"
        .Fill.ForeColor.RGB = RGB(237, 125, 49)
        .Line.Color.RGB = RGB(192, 0, 0)
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = 2
        .TextFrame.VerticalAlignment = 2
        .OnAction = "InventurErstellen"
    End With

    MsgBox "Button an Position E1 erstellt!", vbInformation
End Sub
