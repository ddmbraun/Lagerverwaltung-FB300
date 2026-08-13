Attribute VB_Name = "DuplikatPruefer"
Option Explicit

' ================================================================
'  DUPLIKATE PRUEFEN
'  Erstellt Blatt "Duplikate" mit allen doppelten Artikeln.
'  Duplikat = gleiche ARTIKELNR (rot) ODER gleicher NAME (orange)
'  Aufruf: Alt+F8 -> DuplikatPruefer.Duplikate_Pruefen
' ================================================================
Sub Duplikate_Pruefen()

    ' --- Artikel-Sheet finden ---
    Dim wsA As Worksheet
    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets("Artikel")
    On Error GoTo 0
    If wsA Is Nothing Then
        ' Fallback: Sheet mit "Artikel" im Namen suchen
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Sheets
            If InStr(1, ws.Name, "Artikel", vbTextCompare) > 0 And _
               InStr(1, ws.Name, "Detail", vbTextCompare) = 0 Then
                Set wsA = ws: Exit For
            End If
        Next ws
    End If
    If wsA Is Nothing Then
        MsgBox "Artikel-Sheet nicht gefunden!", vbCritical: Exit Sub
    End If

    ' --- Alle Zeilen einblenden ---
    On Error Resume Next
    wsA.Rows("3:50000").Hidden = False
    wsA.ShowAllData
    On Error GoTo 0

    ' --- Spalten ermitteln (Header in Zeile 2) ---
    Dim cNr As Long, cArt As Long, cEAN As Long
    Dim cLag As Long, cAnz As Long, cVK As Long, cLief As Long
    Dim lastCol As Long
    lastCol = wsA.Cells(2, wsA.Columns.count).End(xlToLeft).Column
    Dim i As Long
    For i = 1 To lastCol
        Dim h As String: h = UCase(Trim(CStr(wsA.Cells(2, i).Value)))
        If h = "ARTIKELNR" Then cNr = i
        If h = "ARTIKEL" Then cArt = i
        If h = "EAN13" Then cEAN = i
        If h = "LAGERORT" Then cLag = i
        If h = "ANZAHL" Then cAnz = i
        If InStr(h, "VK-PREIS") > 0 Then cVK = i
        If h = "LIEFERANT" Then cLief = i
    Next i
    If cArt = 0 Then MsgBox "Spalte ARTIKEL nicht gefunden!", vbCritical: Exit Sub

    ' --- Artikel einlesen ---
    Dim lastRow As Long
    lastRow = wsA.Cells(wsA.Rows.count, cArt).End(xlUp).Row
    If lastRow < 3 Then MsgBox "Keine Artikel gefunden.", vbInformation: Exit Sub

    Dim dictNr   As Object: Set dictNr = CreateObject("Scripting.Dictionary")
    Dim dictName As Object: Set dictName = CreateObject("Scripting.Dictionary")

    Dim anzArtikel As Long: anzArtikel = lastRow - 2
    Dim dupTyp()   As String: ReDim dupTyp(3 To lastRow)
    Dim dupFirst() As Long: ReDim dupFirst(3 To lastRow)
    Dim anzDup     As Long: anzDup = 0

    For i = 3 To lastRow
        Dim artName As String: artName = Trim(CStr(wsA.Cells(i, cArt).Value))
        Dim artNr   As String: artNr = ""
        If cNr > 0 Then artNr = Trim(CStr(wsA.Cells(i, cNr).Value))
        If artName = "" Then GoTo Weiter

        Dim kName As String: kName = LCase(artName)
        Dim kNr   As String: kNr = LCase(artNr)
        Dim isDupNr   As Boolean: isDupNr = False
        Dim isDupName As Boolean: isDupName = False

        If artNr <> "" Then
            If dictNr.Exists(kNr) Then
                isDupNr = True
                dupFirst(i) = dictNr(kNr)
            Else
                dictNr.Add kNr, i
            End If
        End If

        If dictName.Exists(kName) Then
            isDupName = True
            If dupFirst(i) = 0 Then dupFirst(i) = dictName(kName)
        Else
            dictName.Add kName, i
        End If

        If isDupNr And isDupName Then
            dupTyp(i) = "BEIDE": anzDup = anzDup + 1
        ElseIf isDupNr Then
            dupTyp(i) = "NR": anzDup = anzDup + 1
        ElseIf isDupName Then
            dupTyp(i) = "NAME": anzDup = anzDup + 1
        End If
Weiter:
    Next i

    If anzDup = 0 Then
        MsgBox "Keine Duplikate! Alle " & anzArtikel & " Artikel sind eindeutig.", _
               vbInformation, "Duplikat-Prüfung"
        Exit Sub
    End If

    ' --- Ergebnis-Sheet anlegen ---
    Application.DisplayAlerts = False
    On Error Resume Next: ThisWorkbook.Sheets("Duplikate").Delete: On Error GoTo 0
    Application.DisplayAlerts = True

    Dim wsD As Worksheet
    Set wsD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsD.Name = "Duplikate"

    Dim cBlau   As Long: cBlau = RGB(31, 56, 100)
    Dim cWeiss  As Long: cWeiss = RGB(255, 255, 255)
    Dim cGelb   As Long: cGelb = RGB(255, 230, 130)

    ' Titel
    wsD.Range("A1:H1").Merge
    wsD.Cells(1, 1).Value = "DUPLIKAT-PRÜFUNG  –  " & anzDup & " Duplikate  |  " & Format(Now, "DD.MM.YYYY HH:MM")
    wsD.Cells(1, 1).Interior.Color = cBlau
    wsD.Cells(1, 1).Font.Color = cWeiss
    wsD.Cells(1, 1).Font.Bold = True
    wsD.Cells(1, 1).Font.Size = 12
    wsD.Rows(1).RowHeight = 26

    ' Legende Zeile 2
    wsD.Cells(2, 1).Value = "GELB = Original (erster Treffer)"
    wsD.Cells(2, 1).Interior.Color = cGelb
    wsD.Cells(2, 1).Font.Bold = True
    wsD.Cells(2, 2).Value = "ROT = Duplikat gleiche ArtNr"
    wsD.Cells(2, 2).Interior.Color = RGB(255, 160, 160)
    wsD.Cells(2, 2).Font.Bold = True
    wsD.Cells(2, 3).Value = "ORANGE = Duplikat gleicher Name"
    wsD.Cells(2, 3).Interior.Color = RGB(255, 200, 120)
    wsD.Cells(2, 3).Font.Bold = True
    wsD.Rows(2).RowHeight = 20

    ' Header Zeile 3
    Dim hdrs As Variant
    hdrs = Array("Zeile", "Artikelnummer", "Artikelname", "EAN13", "Lagerort", "Bestand", "VK-Preis", "Lieferant")
    Dim c As Integer
    For c = 0 To 7
        wsD.Cells(3, c + 1).Value = hdrs(c)
        wsD.Cells(3, c + 1).Interior.Color = RGB(46, 80, 144)
        wsD.Cells(3, c + 1).Font.Color = cWeiss
        wsD.Cells(3, c + 1).Font.Bold = True
        wsD.Cells(3, c + 1).HorizontalAlignment = xlCenter
    Next c
    wsD.Rows(3).RowHeight = 22

    ' --- Daten schreiben ---
    Dim dRow As Long: dRow = 4
    Dim prevGroup As Long: prevGroup = 0

    For i = 3 To lastRow
        If dupTyp(i) <> "" Then

            ' Neue Gruppe: Trennlinie + Original-Zeile
            If dupFirst(i) <> prevGroup Then
                ' Trennlinie
                If prevGroup > 0 Then
                    wsD.Rows(dRow).RowHeight = 8
                    wsD.Range("A" & dRow & ":H" & dRow).Interior.Color = RGB(210, 210, 210)
                    dRow = dRow + 1
                End If
                ' Original
                Dim orig As Long: orig = dupFirst(i)
                Call ZeileSchreiben(wsD, wsA, dRow, orig, cNr, cArt, cEAN, cLag, cAnz, cVK, cLief, cGelb)
                dRow = dRow + 1
                prevGroup = dupFirst(i)
            End If

            ' Duplikat
            Dim fc As Long
            Select Case dupTyp(i)
                Case "NR": fc = RGB(255, 160, 160)
                Case "NAME": fc = RGB(255, 200, 120)
                Case "BEIDE": fc = RGB(255, 120, 120)
            End Select
            Call ZeileSchreiben(wsD, wsA, dRow, i, cNr, cArt, cEAN, cLag, cAnz, cVK, cLief, fc)
            dRow = dRow + 1
        End If
    Next i

    ' Spaltenbreiten
    wsD.Columns(1).ColumnWidth = 8
    wsD.Columns(2).ColumnWidth = 18
    wsD.Columns(3).ColumnWidth = 45
    wsD.Columns(4).ColumnWidth = 16
    wsD.Columns(5).ColumnWidth = 16
    wsD.Columns(6).ColumnWidth = 10
    wsD.Columns(7).ColumnWidth = 11
    wsD.Columns(8).ColumnWidth = 20

    wsD.Activate
    wsD.Cells(4, 1).Select

    MsgBox anzDup & " Duplikate gefunden (von " & anzArtikel & " Artikeln)." & Chr(10) & Chr(10) & _
           "Zeile A = Zeile im Artikel-Sheet." & Chr(10) & _
           "Entscheide welche Zeilen du löschen möchtest.", _
           vbInformation, "Duplikat-Prüfung"
End Sub

' --- Hilfsfunktion: eine Zeile in Duplikate-Sheet schreiben ---
Private Sub ZeileSchreiben(wsD As Worksheet, wsA As Worksheet, _
    dRow As Long, srcRow As Long, _
    cNr As Long, cArt As Long, cEAN As Long, _
    cLag As Long, cAnz As Long, cVK As Long, cLief As Long, _
    fillColor As Long)

    wsD.Cells(dRow, 1).Value = srcRow
    wsD.Cells(dRow, 3).Value = wsA.Cells(srcRow, cArt).Value
    If cNr > 0 Then wsD.Cells(dRow, 2).Value = CStr(wsA.Cells(srcRow, cNr).Value)
    If cEAN > 0 Then wsD.Cells(dRow, 4).Value = CStr(wsA.Cells(srcRow, cEAN).Value)
    If cLag > 0 Then wsD.Cells(dRow, 5).Value = wsA.Cells(srcRow, cLag).Value
    If cAnz > 0 Then wsD.Cells(dRow, 6).Value = wsA.Cells(srcRow, cAnz).Value
    If cVK > 0 Then
        wsD.Cells(dRow, 7).Value = wsA.Cells(srcRow, cVK).Value
        wsD.Cells(dRow, 7).NumberFormat = "0.00"
    End If
    If cLief > 0 Then wsD.Cells(dRow, 8).Value = wsA.Cells(srcRow, cLief).Value
    wsD.Range("A" & dRow & ":H" & dRow).Interior.Color = fillColor
    wsD.Rows(dRow).RowHeight = 20
End Sub
