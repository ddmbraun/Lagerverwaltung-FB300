' ============================================================
' ERSETZT die vorhandenen GitHub_Export und ExportLagerJSON
' Subs in deinem Modul "NeueModule"
' ============================================================

Sub GitHub_Export()

    On Error GoTo Fehler

    Const GIT_DIR     As String = "D:\_KI-Projekte-2026\07_Lagerverwaltung-Excel\Lagerverw. FB300\"
    Const LAGER_SHEET As String = "Schnellansicht"

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


Private Sub ExportLagerJSON(targetDir As String, sheetName As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Spalten – Zeile 1 = Titel, Zeile 2 = Ueberschriften, Daten ab Zeile 3
    ' Spalte A = leer/intern, Daten beginnen in Spalte B
    Const C_NR          As Integer = 2   ' #
    Const C_ARTNR       As Integer = 3   ' Art.-Nr.
    Const C_ARTIKEL     As Integer = 4   ' Artikel
    Const C_EAN         As Integer = 5   ' EAN
    Const C_VK          As Integer = 6   ' VK-Preis
    Const C_EK          As Integer = 7   ' EK-Preis
    Const C_BESTAND     As Integer = 8   ' Bestand
    Const C_EINHEIT     As Integer = 9   ' Einheit
    Const C_LAGERORT    As Integer = 10  ' Lagerort
    Const C_WARENGRUPPE As Integer = 11  ' Warengruppe
    Const C_ATTRIBUT    As Integer = 12  ' Attribut

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, C_ARTNR).End(xlUp).Row

    Dim json    As String
    Dim isFirst As Boolean
    isFirst = True
    json = "[" & vbNewLine

    Dim i     As Long
    Dim artNr As String

    For i = 3 To lastRow   ' Zeile 1 = Titel, Zeile 2 = Ueberschriften

        artNr = Trim(CStr(ws.Cells(i, C_ARTNR).Value))
        If artNr = "" Then GoTo NextRow

        If Not isFirst Then json = json & "," & vbNewLine
        isFirst = False

        json = json & "  {" & _
            """nr"":"           & JStr(ws.Cells(i, C_NR))          & "," & _
            """artnr"":"        & JStr(ws.Cells(i, C_ARTNR))       & "," & _
            """artikel"":"      & JStr(ws.Cells(i, C_ARTIKEL))     & "," & _
            """ean"":"          & JStr(ws.Cells(i, C_EAN))         & "," & _
            """vk"":"           & JStr(ws.Cells(i, C_VK))          & "," & _
            """ek"":"           & JStr(ws.Cells(i, C_EK))          & "," & _
            """bestand"":"      & JStr(ws.Cells(i, C_BESTAND))     & "," & _
            """einheit"":"      & JStr(ws.Cells(i, C_EINHEIT))     & "," & _
            """lagerort"":"     & JStr(ws.Cells(i, C_LAGERORT))    & "," & _
            """warengruppe"":"  & JStr(ws.Cells(i, C_WARENGRUPPE)) & "," & _
            """attribut"":"     & JStr(ws.Cells(i, C_ATTRIBUT))    & _
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


Private Function JStr(cell As Range) As String
    Dim s As String
    s = CStr(cell.Value)
    s = Replace(s, "\",  "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr,   " ")
    s = Replace(s, vbLf,   " ")
    JStr = """" & s & """"
End Function
