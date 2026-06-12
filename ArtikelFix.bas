Attribute VB_Name = "ArtikelFix"
Option Explicit

' ================================================================
'  ARTIKEL-SHEET FIXES
'  Reihenfolge:
'    1. ArtikelFix.Layout_Neu_Einrichten   -> A1=Spacer, B1=Suche, D1=Treffer
'    2. ArtikelFix.Events_Neu_Installieren -> Events auf B1 umstellen
' ================================================================


' ----------------------------------------------------------------
'  1) Buttons vergroessern (einmalig)
' ----------------------------------------------------------------
Sub Buttons_Groesser()
    Dim wsA As Worksheet
    Set wsA = ArtikelSheet()
    If wsA Is Nothing Then Exit Sub

    Dim neueHoehe As Double: neueHoehe = 30
    Dim neueZeile As Double: neueZeile = 38

    Application.ScreenUpdating = False
    wsA.Rows(1).RowHeight = neueZeile

    Dim shp As Shape
    For Each shp In wsA.Shapes
        If shp.Top < neueZeile + 4 And shp.Name <> "lbl_Treffer" Then
            shp.Height = neueHoehe
            shp.Top = (neueZeile - neueHoehe) / 2
            shp.TextFrame.Characters.Font.Size = 10
        End If
    Next shp

    Application.ScreenUpdating = True
    MsgBox "Buttons vergroessert!", vbInformation
End Sub


' ----------------------------------------------------------------
'  2) Layout neu einrichten (EINMALIG):
'     A1  = leeres Spacer-Feld
'     B1:C1 = Suchfeld (merged)
'     D1  = Treffer-Anzeige (dezent)
'     btn_FilterX wird vor D1 positioniert
' ----------------------------------------------------------------
Sub Layout_Neu_Einrichten()
    Dim wsA As Worksheet
    Set wsA = ArtikelSheet()
    If wsA Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Alte Merges aufloesen
    On Error Resume Next
    wsA.Range("A1:D1").UnMerge
    wsA.Range("A1:C1").UnMerge
    wsA.Range("B1:C1").UnMerge
    On Error GoTo 0

    ' A1: Spacer (leer, neutrale Farbe)
    With wsA.Range("A1")
        .Value = ""
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlNone
    End With

    ' B1:C1: Suchfeld
    wsA.Range("B1:C1").Merge
    With wsA.Range("B1")
        .Value = "  Suche..."
        .Font.Color = RGB(150, 150, 150)
        .Font.Italic = True
        .Font.Size = 11
        .Font.Bold = False
        .Interior.Color = RGB(250, 250, 250)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' D1: Treffer-Anzeige (dezent, kein Button-Look)
    With wsA.Range("D1")
        .Value = ""
        .Interior.Color = RGB(235, 235, 235)
        .Font.Color = RGB(50, 50, 50)
        .Font.Bold = True
        .Font.Italic = False
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(180, 180, 180)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(180, 180, 180)
    End With

    ' btn_FilterX vor D1 verschieben (am Ende von C1)
    Dim btnX As Shape
    On Error Resume Next
    Set btnX = wsA.Shapes("btn_FilterX")
    On Error GoTo 0
    If Not btnX Is Nothing Then
        btnX.Left = wsA.Range("D1").Left - btnX.Width - 3
        btnX.Top = (wsA.Rows(1).RowHeight - btnX.Height) / 2
    End If

    ' Altes Treffer-Shape loeschen
    On Error Resume Next
    wsA.Shapes("lbl_Treffer").Delete
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "Layout fertig!" & Chr(10) & "Jetzt ArtikelFix.Events_Neu_Installieren ausfuehren!", vbInformation
End Sub


' ----------------------------------------------------------------
'  3) Tabelle9 Events neu installieren (Suchfeld = B1)
' ----------------------------------------------------------------
Sub Events_Neu_Installieren()
    Dim wsA As Worksheet
    Set wsA = ArtikelSheet()
    If wsA Is Nothing Then Exit Sub

    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsA.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim n As String: n = Chr(10)
    Dim c As String: c = ""

    ' SelectionChange: Platzhalter in B1
    c = c & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & n
    c = c & "    On Error Resume Next" & n
    c = c & "    If Not Intersect(Target, Me.Range(""B1"")) Is Nothing Then" & n
    c = c & "        If Trim(Me.Range(""B1"").Value) = ""Suche..."" Or Me.Range(""B1"").Value = """" Then" & n
    c = c & "            Application.EnableEvents = False" & n
    c = c & "            Me.Range(""B1"").Value = """"" & n
    c = c & "            Me.Range(""B1"").Font.Color = RGB(0, 0, 0)" & n
    c = c & "            Me.Range(""B1"").Font.Italic = False" & n
    c = c & "            Application.EnableEvents = True" & n
    c = c & "        End If" & n
    c = c & "        Application.SendKeys ""{F2}""" & n
    c = c & "    Else" & n
    c = c & "        If Trim(Me.Range(""B1"").Value) = """" Then" & n
    c = c & "            Application.EnableEvents = False" & n
    c = c & "            Me.Range(""B1"").Value = ""  Suche...""" & n
    c = c & "            Me.Range(""B1"").Font.Color = RGB(150, 150, 150)" & n
    c = c & "            Me.Range(""B1"").Font.Italic = True" & n
    c = c & "            Application.EnableEvents = True" & n
    c = c & "        End If" & n
    c = c & "    End If" & n
    c = c & "End Sub" & n
    c = c & n

    ' Worksheet_Change: Suche aus B1
    c = c & "Private Sub Worksheet_Change(ByVal Target As Range)" & n
    c = c & "    If Not Intersect(Target, Me.Range(""B1"")) Is Nothing Then" & n
    c = c & "        Dim s As String: s = Trim(CStr(Me.Range(""B1"").Value))" & n
    c = c & "        Application.EnableEvents = False" & n
    c = c & "        If s = """" Or s = ""Suche..."" Then" & n
    c = c & "            ArtikelFix.Filter_Loeschen_Fix Me" & n
    c = c & "        Else" & n
    c = c & "            ArtikelFix.Artikel_Suchen_Fix Me, s" & n
    c = c & "        End If" & n
    c = c & "        Application.EnableEvents = True" & n
    c = c & "    End If" & n
    c = c & "End Sub" & n

    cm.AddFromString c
    MsgBox "Events installiert! Suche in B1 testen.", vbInformation
End Sub


' ----------------------------------------------------------------
'  4) Suche ausfuehren + Treffer in D1
' ----------------------------------------------------------------
Sub Artikel_Suchen_Fix(wsA As Worksheet, such As String)
    Application.ScreenUpdating = False

    wsA.Rows("3:50000").Hidden = False
    On Error Resume Next
    wsA.ShowAllData
    On Error GoTo 0

    Dim cArt As Long: cArt = LagerMakros.Spalte_Finden(wsA, "ARTIKEL")
    Dim cEAN As Long: cEAN = LagerMakros.Spalte_Finden(wsA, "EAN13")
    If cArt = 0 Then GoTo Fertig

    Dim lastRow As Long
    lastRow = wsA.Cells(wsA.Rows.count, cArt).End(xlUp).Row
    If lastRow < 3 Then GoTo Fertig

    Dim nurEAN As Boolean: nurEAN = (Len(such) >= 8 And IsNumeric(such))
    Dim woerter() As String: woerter = Split(LCase(such), " ")

    Dim treffer As Long: treffer = 0
    Dim i As Long
    For i = 3 To lastRow
        Dim artName As String: artName = LCase(CStr(wsA.Cells(i, cArt).Value))
        Dim match As Boolean: match = False

        If nurEAN Then
            If cEAN > 0 Then match = (CStr(wsA.Cells(i, cEAN).Value) = such)
        Else
            Dim w As Variant
            match = True
            For Each w In woerter
                If Trim(CStr(w)) <> "" Then
                    If InStr(artName, CStr(w)) = 0 Then match = False: Exit For
                End If
            Next w
        End If

        If match Then
            treffer = treffer + 1
        Else
            wsA.Rows(i).Hidden = True
        End If
    Next i

Fertig:
    On Error Resume Next
    wsA.Range("D1").Value = treffer & " Treffer"
    wsA.Shapes("lbl_Treffer").Delete
    On Error GoTo 0

    Application.ScreenUpdating = True
End Sub


' ----------------------------------------------------------------
'  5a) Filter loeschen – aufrufbar per Alt+F8 (kein Parameter)
' ----------------------------------------------------------------
Sub Alle_Artikel_Anzeigen()
    Dim wsA As Worksheet
    Set wsA = ArtikelSheet()
    If wsA Is Nothing Then Exit Sub
    Filter_Loeschen_Fix wsA
    MsgBox "Alle Artikel angezeigt.", vbInformation
End Sub


' ----------------------------------------------------------------
'  5b) X-Button wiederherstellen (falls verschwunden)
' ----------------------------------------------------------------
Sub FilterX_Wiederherstellen()
    Dim wsA As Worksheet
    Set wsA = ArtikelSheet()
    If wsA Is Nothing Then Exit Sub

    ' Alten X-Button entfernen falls vorhanden
    On Error Resume Next
    wsA.Shapes("btn_FilterX").Delete
    On Error GoTo 0

    ' Neu erstellen: rechts neben C1 (vor D1)
    Dim xLeft As Double: xLeft = wsA.Range("D1").Left - 26
    Dim xTop  As Double: xTop = (wsA.Rows(1).RowHeight - 20) / 2
    Dim shp As Shape
    Set shp = wsA.Shapes.AddShape(msoShapeRoundedRectangle, xLeft, xTop, 22, 20)
    With shp
        .Name = "btn_FilterX"
        .Fill.ForeColor.RGB = RGB(200, 50, 50)
        .Line.Visible = msoFalse
        .Placement = xlMoveAndSize
    End With
    With shp.TextFrame
        .Characters.Text = "X"
        .Characters.Font.Color = RGB(255, 255, 255)
        .Characters.Font.Bold = True
        .Characters.Font.Size = 9
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With

    ' Makro zuweisen
    shp.OnAction = "ArtikelFix.Alle_Artikel_Anzeigen"

    MsgBox "X-Button wiederhergestellt!", vbInformation
End Sub


' ----------------------------------------------------------------
'  5) Filter loeschen (
