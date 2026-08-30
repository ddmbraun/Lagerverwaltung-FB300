Attribute VB_Name = "LagerMakros"

' Einstellungen
Const ZEBRA_DRUCKER   As String = "ZDesigner GK420d"
Public Const BENUTZER        As String = "Frank"
Const INV_DATEN_START As Long = 6       ' Inventur: Datenzeilen ab Zeile 6

' Artikel-Sheet: Zeile 1=Titel, 2=Suchfeld, 3=Toolbar, 4=Ueberschriften -> Daten ab 5.
' Wer hier mit 3 startet, liest die Ueberschriftenzeile als Artikel ein.
Const ART_DATEN_START As Long = 5

' ================================================================
'  SCHNELLANSICHT - das Spaltenbild steht NUR hier
' ================================================================
' Bis zum 29.08.2026 gab es auf diesem Blatt ZWEI Layouts nebeneinander:
' die Ueberschriften, der EK-Knopf und der Detail-Doppelklick erwarteten
' 11 Spalten, die Fuell-Routine schrieb aber ein altes 6-Spalten-Bild
' darueber. Ergebnis: die Spalten G-M blieben als Altdaten stehen und
' gehoerten zur NAECHSTEN Zeile, und das Such-Fenster zeigte den EK-Preis
' als "Bestand" an.
' Deshalb stehen die Spaltennummern jetzt an EINER Stelle. Wer das Bild
' aendert, aendert es hier - und Schnellansicht_Aktualisieren schreibt die
' Ueberschriften bei jedem Lauf mit, damit Kopf und Daten nicht wieder
' auseinanderlaufen koennen.
Const SV_KOPF_ZEILE   As Long = 3       ' Ueberschriftenzeile
Const SV_DATEN_START  As Long = 4       ' erste Datenzeile
Const SV_TREFFER      As Long = 8       ' H2: Trefferanzeige neben den Knoepfen
Const SV_NR           As Long = 2       ' B: laufende Nummer
Const SV_ARTNR        As Long = 3       ' C: Artikelnummer
Const SV_ARTIKEL      As Long = 4       ' D: Artikelname   (danach wird gesucht)
Const SV_EAN          As Long = 5       ' E: EAN13         (danach wird gesucht)
Const SV_VK           As Long = 6       ' F: VK-Preis
Const SV_EK           As Long = 7       ' G: EK-Preis      (den blendet "EK ausbl." aus)
Const SV_BESTAND      As Long = 8       ' H: Bestand
Const SV_EINHEIT      As Long = 9       ' I: Einheit
Const SV_LAGERORT     As Long = 10      ' J: Lagerort
Const SV_WARENGRUPPE  As Long = 11      ' K: Warengruppe
Const SV_ATTRIBUT     As Long = 12      ' L: Attribut
Const SV_ZEILENVERWEIS As Long = 13     ' M: Zeile im Artikel-Blatt (versteckt)
Const SV_SPALTEN      As Long = 13      ' so breit ist die Liste insgesamt

' Die Knoepfe in Zeile 2 - Doppelklick darauf. ⛔ Sie stehen NICHT dicht
' beieinander: zwischen "EK ausbl." (G) und "SCHLIESSEN" (J) liegen die
' Trefferanzeige (H) und eine leere Spalte (I). Ein Bereich "4 bis 8" wie
' frueher trifft SCHLIESSEN deshalb nie und dafuer die Trefferanzeige.
Const SV_BTN_SUCHEN     As Long = 4     ' D2  SUCHEN
Const SV_BTN_FILTER     As Long = 5     ' E2  FILTER LOESCHEN
Const SV_BTN_AKTUELL    As Long = 6     ' F2  AKTUALISIEREN
Const SV_BTN_EK         As Long = 7     ' G2  EK ausbl. / EK einbl.
Const SV_BTN_SCHLIESSEN As Long = 10    ' J2  SCHLIESSEN

' Kennwort fuer den Blattschutz waehrend der Inventur.
' � Kein echter Schutz - es steht hier im Klartext und Excel-Blattschutz laesst sich
' ohnehin aushebeln. Zweck ist allein, VERSEHENTLICHES Aendern oder Loeschen zu
' verhindern, wenn jemand anderes zaehlt.
Const INV_SCHUTZ As String = "inventur2026"

' Wieviele Treffer hoechstens angeboten werden. Darueber lohnt das Auswaehlen nicht,
' dann ist "genauer suchen" schneller.
Const INV_MAX_TREFFER As Long = 20

' Zaehlblatt: Trefferliste und Kontrollstreifen teilen sich denselben Bereich
Const ZB_LISTE_START As Long = 15       ' erste Zeile darunter

' Zaehlblatt: Spalten A und B bleiben als linker Rand frei, das Formular beginnt bei C.
Const ZB_LABEL As Long = 3              ' C - Beschriftungen
Const ZB_WERT  As Long = 4              ' D - Werte (Art.-Nr., SOLL)
Const ZB_L2    As Long = 6              ' F - zweite Beschriftungsspalte (EAN, MENGE)
Const ZB_W2    As Long = 7              ' G - Mengenfeld
Const ZB_ENDE  As Long = 9              ' I - rechter Rand des Formulars
' Platzhalter im leeren aktiven Feld - sieht aus wie ein wartender Cursor.
' ACHTUNG: Diese Texte duerfen NIE als Suchbegriff oder Menge durchgehen (siehe
' Zaehlen_Suchen und Zaehlen_Eintragen).
Const ZB_PLATZ_SUCHE As String = "|   hier scannen ..."
Const ZB_PLATZ_MENGE As String = "|"

Const ZB_SUCHE   As String = "D5:H5"    ' Suchfeld (verbunden)
Const ZB_ARTIKEL As String = "D7:I7"    ' Artikelname (verbunden)
Const ZB_EAN     As String = "G8:H8"    ' EAN (verbunden)
Const ZB_LAGER   As String = "D9:E9"    ' Lagerort (verbunden)

' Inventurliste - Spalten hinter der sichtbaren Liste (A-J)
Const INV_COL_DATUM  As Long = 11       ' K: gezaehlt am
Const INV_COL_WER    As Long = 12       ' L: gezaehlt von
Const INV_COL_ZEILE  As Long = 13       ' M: Artikelzeile im Artikel-Sheet (versteckt)
Const INV_COL_VORHER As Long = 14       ' N: Bestand vor der Uebernahme (versteckt)
Const INV_COL_LETZTE As Long = 14       ' letzte belegte Spalte der Liste

' Letzte angeklickte Artikelzeile
Public g_LetzteZeile As Long
' Gewaehlter Artikel in InvSuche
Public g_InvSucheArtikelZeile As Long
' Gewaehlter Artikel in der Inventurliste (aus Inventur_Suchen)
Public g_InvArtikelZeile As Long

' ================================================================
'  LETZTE BELEGTE ZEILE - unabhaengig von Filtern
' ================================================================
' ⛔ End(xlUp) liefert bei aktivem Filter die letzte SICHTBARE Zeile, nicht die letzte
' belegte. Am 16.08.2026 meldete das Befuellen dadurch 5.683 statt 6.986 Artikel:
' im Artikel-Blatt war noch eine Suche aktiv, die Zeilen ausgeblendet hatte.
' 1.303 Artikel waeren stillschweigend aus der Inventur gefallen.
' Find(SearchDirection:=xlPrevious) sieht auch ausgeblendete Zeilen.
Function LetzteZeile(ws As Worksheet, spalte As Long) As Long
    On Error Resume Next
    LetzteZeile = 0
    Dim f As Range
    Set f = ws.Columns(spalte).Find(What:="*", LookIn:=xlFormulas, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then LetzteZeile = f.Row
    ' Sicherheitsnetz: falls Find nichts findet, der alte Weg
    If LetzteZeile = 0 Then LetzteZeile = ws.Cells(ws.Rows.count, spalte).End(xlUp).Row
    Err.Clear
End Function

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
    Dim ws As Worksheet: Set ws = Target.Worksheet
    If Target.Row >= 3 Then
        ' Vorherige Markierung entfernen
        If g_LetzteZeile >= 3 Then
            ws.Rows(g_LetzteZeile).Interior.ColorIndex = xlNone
        End If
        ' Neue Zeile hellgelb markieren
        g_LetzteZeile = Target.Row
        ws.Rows(g_LetzteZeile).Interior.Color = 16776960
        Exit Sub
    End If
    If Target.Row = 1 Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Select Case Target.Column
            Case 4, 5: ZuAbgang_Buchen
            Case 6, 7: Etikett_Drucken
            Case 8, 9: EK_Toggle
            Case 10, 11: Filter_Loeschen
            Case 12, 13: NeueModule.Schnellansicht_Oeffnen
            Case 14, 15: NeuerArtikel
        End Select
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

' ================================================================
'  ZU-/ABGANG BUCHEN
' ================================================================
Sub ZuAbgang_Buchen()
    If g_LetzteZeile < 3 Then
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
    Dim aktuell  As Double: aktuell = val(wsA.Cells(zeile, colAnz).Value)

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
    Dim menge As Double: menge = val(eingabe)
    If menge = 0 Then MsgBox "Ungueltige Eingabe.": Exit Sub

    Dim typ As String: typ = IIf(menge > 0, "Zugang", "Abgang")
    Dim neuerBestand As Double: neuerBestand = aktuell + menge

    ' Artikel-Sheet aktualisieren
    wsA.Cells(zeile, colAnz).Value = neuerBestand

    ' Bewegung eintragen
    Dim nRow As Long
    nRow = wsZ.Cells(wsZ.Rows.count, 1).End(xlUp).Row + 1
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
    For i = 2 To wsB.Cells(wsB.Rows.count, 2).End(xlUp).Row
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
    If g_LetzteZeile < 3 Then
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
    wsA.Cells(1, 8).Value = IIf(wsA.Columns(ekSpalte).Hidden, "EK einblenden", "EK ausblenden")
End Sub

' ================================================================
'  FILTER LOESCHEN
' ================================================================
Sub Filter_Loeschen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    ' Markierung entfernen
    If g_LetzteZeile >= 5 Then
        wsA.Rows(g_LetzteZeile).Interior.ColorIndex = xlNone
    End If
    If wsA.AutoFilterMode Then wsA.AutoFilterMode = False
    wsA.Rows("5:50000").Hidden = False
    Application.EnableEvents = False
    wsA.Cells(2, 2).Value = ""
    Application.EnableEvents = True
    g_LetzteZeile = 0
    Artikel_Anzahl_Anzeigen
End Sub

' ================================================================
'  HILFSFUNKTIONEN
' ================================================================
Function Spalte_Finden(ws As Worksheet, headerName As String) As Long
    ' Artikel-Sheet: Zeile 1=Titel, 2=Suchfeld -> ab Zeile 3 suchen (Buttons/Header)
    ' Alle anderen Sheets: ab Zeile 1 suchen
    Dim startRow As Long
    startRow = IIf(InStr(ws.Name, "rtikel") > 0, 3, 1)

    Dim i As Long, r As Long, lastCol As Long
    For r = startRow To startRow + 5
        lastCol = ws.Cells(r, ws.Columns.count).End(xlToLeft).Column
        For i = 1 To lastCol
            If InStr(1, CStr(ws.Cells(r, i).Value), headerName, vbTextCompare) > 0 Then
                Spalte_Finden = i
                Exit Function
            End If
        Next i
    Next r

    Spalte_Finden = 0
End Function

' ================================================================
'  SCHNELLANSICHT - HANDLER (Buttons in Zeile 2)
' ================================================================
' Rueckgabe TRUE = es war ein Knopf und er wurde ausgefuehrt. Der Blattcode
' setzt daraufhin Cancel = True. ⛔ Wichtig, dass NICHT pauschal fuer die ganze
' Zeile 2 abgebrochen wird: in B2 steht das Suchfeld, dort muss ein Doppelklick
' weiterhin den Schreibcursor setzen.
Function Schnellansicht_Handler(ByVal Target As Range) As Boolean
    Schnellansicht_Handler = False
    If Target.Row <> 2 Then Exit Function
    Dim col As Long: col = Target.Column
    Select Case col
        Case SV_BTN_SUCHEN, SV_BTN_FILTER, SV_BTN_AKTUELL, SV_BTN_EK, SV_BTN_SCHLIESSEN
            Schnellansicht_Handler = True
        Case Else
            Exit Function
    End Select

    On Error GoTo Fehler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Select Case col
        Case SV_BTN_SUCHEN:     Schnellansicht_Suchen
        Case SV_BTN_FILTER:     Schnellansicht_FilterLoeschen
        Case SV_BTN_AKTUELL:    Schnellansicht_Aktualisieren
        Case SV_BTN_EK:         NeueModule.Schnellansicht_EK_Toggle
        Case SV_BTN_SCHLIESSEN: NeueModule.Schnellansicht_Schliessen
    End Select
Fehler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function


' ================================================================
'  SCHNELLANSICHT - Blattcode einsetzen (Knoepfe + Detail-Doppelklick)
' ================================================================
' ⛔ Dieser Text steht NUR HIER. Setup_Ausfuehren ruft dieselbe Routine auf,
'   statt ihn ein zweites Mal zusammenzubauen.
' Braucht Zugriff auf das VBA-Projekt:
'   Datei > Optionen > Trust Center > Einstellungen fuer das Trust Center >
'   Makroeinstellungen > "Zugriff auf das VBA-Projektobjektmodell vertrauen"
Private Function Schnellansicht_BlattcodeSchreiben() As Boolean
    Schnellansicht_BlattcodeSchreiben = False

    ' ⛔ Die Zahlen im erzeugten Text stehen fest und duerfen NICHT mit "&" aus
    '   den Konstanten zusammengesetzt werden - dann sieht _pruefung\event_pruefen.py
    '   den Blattcode nicht mehr (die liest nur die Zeichenketten). Damit die
    '   Zahlen trotzdem nicht auseinanderlaufen, hier die Wache davor:
    ' ⚠ Bedingung bewusst auf EINER Zeile: _pruefung\vba_pruefen.py zaehlt
    '   Fortsetzungszeilen nicht zusammen und meldet sonst ein If zu wenig.
    If SV_DATEN_START <> 4 Or SV_NR <> 2 Or SV_ATTRIBUT <> 12 Or SV_ARTIKEL <> 4 Or SV_EAN <> 5 Then
        MsgBox "Das Spaltenbild (SV_-Konstanten) wurde geaendert, der Blattcode" & Chr(10) & _
               "darunter aber nicht. Bitte beides angleichen." & Chr(10) & Chr(10) & _
               "Es wurde nichts geschrieben.", vbCritical, "Schnellansicht"
        Exit Function
    End If

    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then
        MsgBox "Schnellansicht-Blatt nicht gefunden!", vbCritical
        Exit Function
    End If

    On Error GoTo KeinZugriff
    Dim vb As Object: Set vb = ThisWorkbook.VBProject.VBComponents(wsS.CodeName)
    Dim cm As Object: Set cm = vb.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

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
    c = c & "    If Target.Row = 2 Then" & Chr(10)
    c = c & "        If LagerMakros.Schnellansicht_Handler(Target) Then Cancel = True" & Chr(10)
    c = c & "        Exit Sub" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "    If Target.Row >= 4 Then" & Chr(10)
    c = c & "        If Target.Column >= 2 And Target.Column <= 12 Then" & Chr(10)
    c = c & "            Dim sArt As String" & Chr(10)
    c = c & "            sArt = CStr(Me.Cells(Target.Row, 4).Value)" & Chr(10)
    c = c & "            If sArt <> """" Then" & Chr(10)
    c = c & "                Cancel = True" & Chr(10)
    c = c & "                Dim sEAN As String" & Chr(10)
    c = c & "                sEAN = CStr(Me.Cells(Target.Row, 5).Value)" & Chr(10)
    c = c & "                NeueModule.SchnellDetail_Laden sEAN, sArt" & Chr(10)
    c = c & "            End If" & Chr(10)
    c = c & "        End If" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c

    Schnellansicht_BlattcodeSchreiben = True
    Exit Function

KeinZugriff:
    MsgBox "Excel laesst den Zugriff auf den Makro-Code nicht zu." & Chr(10) & Chr(10) & _
           "Bitte einschalten unter:" & Chr(10) & _
           "Datei > Optionen > Trust Center > Einstellungen fuer das Trust Center >" & Chr(10) & _
           "Makroeinstellungen > ""Zugriff auf das VBA-Projektobjektmodell vertrauen""" & Chr(10) & Chr(10) & _
           "Danach Excel neu starten und den Schritt wiederholen." & Chr(10) & Chr(10) & _
           "(Fehler " & Err.Number & ": " & Err.Description & ")", _
           vbExclamation, "Kein Zugriff auf den Makro-Code"
End Function


' ================================================================
'  SCHNELLANSICHT - Knoepfe einrichten (einzeln ueber Alt+F8 aufrufbar)
' ================================================================
' Richtet NUR das Schnellansicht-Blatt ein. Absichtlich getrennt von
' Setup_Ausfuehren: das baut zusaetzlich Toolbar-Shapes und Inventurblaetter
' neu auf - das braucht es hier nicht, und weniger anfassen ist sicherer.
Sub Setup_Schnellansicht_Knoepfe()
    If Not Schnellansicht_BlattcodeSchreiben() Then Exit Sub
    MsgBox "Die Knoepfe der Schnellansicht sind eingerichtet." & Chr(10) & Chr(10) & _
           "Doppelklick wirkt jetzt auf:" & Chr(10) & _
           "  SUCHEN, FILTER LOESCHEN, AKTUALISIEREN, EK ausbl., SCHLIESSEN" & Chr(10) & Chr(10) & _
           "Ein Doppelklick auf eine Artikelzeile oeffnet die Detailansicht." & Chr(10) & Chr(10) & _
           "Bitte die Mappe jetzt speichern (Strg+S).", _
           vbInformation, "Schnellansicht eingerichtet"
End Sub

' ================================================================
'  SCHNELLANSICHT SUCHEN + POPUP (mit Mehrwort-Suche)
' ================================================================
Sub Schnellansicht_Suchen()
    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub

    Dim such As String: such = Trim(CStr(wsS.Cells(2, 2).Value))

    ' Leer = alles anzeigen
    If such = "" Then
        Schnellansicht_FilterLoeschen
        Exit Sub
    End If

    Dim lastSvRow As Long
    lastSvRow = wsS.Cells(wsS.Rows.count, SV_ARTIKEL).End(xlUp).Row
    If lastSvRow < SV_DATEN_START Then
        MsgBox "Die Schnellansicht ist noch leer." & Chr(10) & _
               "Bitte zuerst auf AKTUALISIEREN klicken.", vbExclamation, "Suche"
        Exit Sub
    End If

    ' Suchbegriffe aufteilen (Leerzeichen = UND-Verknuepfung)
    Dim woerter() As String: woerter = Split(LCase(such), " ")

    ' Reine Ziffernfolge? Dann auch EAN und Artikelnummer durchsuchen.
    ' ⛔ Frueher stand hier  such = CStr(val(such))  - das ging bei einer
    '   13-stelligen EAN schief, weil val() daraus 5,90017151051E+12 macht.
    '   Ausgerechnet die EAN-Suche fiel damit durch. Deshalb Ziffer fuer Ziffer.
    Dim nurZiffern As Boolean, zi As Long
    nurZiffern = (Len(such) > 0)
    For zi = 1 To Len(such)
        If InStr("0123456789", Mid$(such, zi, 1)) = 0 Then
            nurZiffern = False
            Exit For
        End If
    Next zi

    Application.ScreenUpdating = False
    On Error GoTo Fehler
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False

    ' Die Liste in EINEM Zug lesen - Zelle fuer Zelle waere bei 7.000 Artikeln zaeh.
    Dim anz As Long: anz = lastSvRow - SV_DATEN_START + 1
    Dim dat As Variant
    dat = wsS.Cells(SV_DATEN_START, 1).Resize(anz, SV_SPALTEN).Value

    Dim sicht() As Boolean: ReDim sicht(1 To anz)
    Dim treffer As Long, gef As Long
    Dim i As Long, w As Integer, passt As Boolean, suchIn As String

    For i = 1 To anz
        If nurZiffern Then
            suchIn = LCase(CStr(dat(i, SV_ARTIKEL)) & " " & _
                           CStr(dat(i, SV_EAN)) & " " & CStr(dat(i, SV_ARTNR)))
        Else
            suchIn = LCase(CStr(dat(i, SV_ARTIKEL)))
        End If
        passt = (Trim(CStr(dat(i, SV_ARTIKEL))) <> "")
        If passt Then
            For w = 0 To UBound(woerter)
                If Trim(woerter(w)) <> "" Then
                    If InStr(suchIn, Trim(woerter(w))) = 0 Then
                        passt = False
                        Exit For
                    End If
                End If
            Next w
        End If
        sicht(i) = passt
        If passt Then
            treffer = treffer + 1
            If treffer = 1 Then gef = i
        End If
    Next i

    wsS.Cells(2, SV_TREFFER).Value = treffer & " Treffer"

    ' --- Kein Treffer: Liste unveraendert lassen ------------------------------
    If treffer = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Kein Artikel gefunden für: """ & such & """", vbExclamation, "Suche"
        Exit Sub
    End If

    ' --- Genau 1 Treffer: alles aus der Liste selbst, kein zweites Nachschlagen
    If treffer = 1 Then
        Application.ScreenUpdating = True
        ' ⛔ Die Werte kommen als Zahl aus dem Blatt. Nicht ueber CStr+val gehen -
        '   aus 5,5 wuerde dabei 5 (val rechnet mit Punkt, nicht mit Komma).
        Dim bestandZahl As Double
        If IsNumeric(dat(gef, SV_BESTAND)) Then bestandZahl = CDbl(dat(gef, SV_BESTAND))
        Dim einheit As String: einheit = Trim(CStr(dat(gef, SV_EINHEIT)))
        Dim bestand As String
        If bestandZahl = 0 Then
            bestand = "0 " & einheit & "   !! NACHBESTELLUNG !!"
        ElseIf bestandZahl <= 5 Then
            bestand = Format(bestandZahl, "0") & " " & einheit & "  (Bestand niedrig!)"
        Else
            bestand = Format(bestandZahl, "0") & " " & einheit
        End If
        MsgBox CStr(dat(gef, SV_ARTIKEL)) & Chr(10) & _
               String(40, "-") & Chr(10) & _
               "EAN:         " & CStr(dat(gef, SV_EAN)) & Chr(10) & _
               "ArtNr:       " & CStr(dat(gef, SV_ARTNR)) & Chr(10) & _
               "VK-Preis:    " & Format(dat(gef, SV_VK), "0.00") & " EUR" & Chr(10) & _
               "Bestand:     " & bestand & Chr(10) & _
               "Lagerort:    " & CStr(dat(gef, SV_LAGERORT)) & Chr(10) & _
               "Warengruppe: " & CStr(dat(gef, SV_WARENGRUPPE)), _
               vbInformation, "Artikel gefunden"
        Exit Sub
    End If

    ' --- Mehrere Treffer: Nicht-Treffer ausblenden ---------------------------
    ' Zeile fuer Zeile waere bei 7.000 Artikeln sehr langsam. Darum erst alles
    ' einblenden und dann zusammenhaengende Bloecke auf einmal ausblenden.
    ' ⚠ Eine Bereichsangabe fuer Range() darf nur 255 Zeichen lang sein -
    '   deshalb wird zwischendurch geleert.
    wsS.Rows(SV_DATEN_START & ":" & lastSvRow).Hidden = False
    Dim bloecke As String, offen As Long
    For i = 1 To anz + 1
        passt = (i > anz)
        If Not passt Then passt = sicht(i)
        If Not passt Then
            If offen = 0 Then offen = i
        ElseIf offen > 0 Then
            If bloecke <> "" Then bloecke = bloecke & ","
            bloecke = bloecke & (offen + SV_DATEN_START - 1) & ":" & (i - 1 + SV_DATEN_START - 1)
            offen = 0
            If Len(bloecke) > 200 Then
                wsS.Range(bloecke).EntireRow.Hidden = True
                bloecke = ""
            End If
        End If
    Next i
    If bloecke <> "" Then wsS.Range(bloecke).EntireRow.Hidden = True

    Application.ScreenUpdating = True
    Exit Sub

Fehler:
    Application.ScreenUpdating = True
    MsgBox "Fehler " & Err.Number & " beim Suchen:" & Chr(10) & Err.Description, _
           vbCritical, "Schnellansicht"
End Sub

' ================================================================
'  SCHNELLANSICHT FILTER LOESCHEN
' ================================================================
Sub Schnellansicht_FilterLoeschen()
    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsS Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False
    ' Die letzte Zeile ueber die Artikelspalte suchen - Spalte 2 traegt nur
    ' die laufende Nummer und waere nach einem Layout-Wechsel wieder falsch.
    Dim lastRow As Long
    lastRow = wsS.Cells(wsS.Rows.count, SV_ARTIKEL).End(xlUp).Row
    If lastRow >= SV_DATEN_START Then
        wsS.Rows(SV_DATEN_START & ":" & lastRow).Hidden = False
    End If
    wsS.Cells(2, 2).Value = ""
    wsS.Cells(2, SV_TREFFER).Value = ""
    Schnellansicht_NachOben wsS
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  SCHNELLANSICHT - zurueck an den Anfang
' ================================================================
' Nach dem Aufraeumen stand man sonst irgendwo mitten in 7.000 Zeilen und
' musste von Hand hochrollen (Frank, 30.08.2026). Der Cursor landet gleich
' im Suchfeld - dann kann man sofort weitertippen.
' ⚠ ActiveWindow gibt es nur fuer das gerade sichtbare Blatt, deshalb die
'   Abfrage. Wird die Routine im Hintergrund gerufen, passiert nichts.
Private Sub Schnellansicht_NachOben(wsS As Worksheet)
    On Error Resume Next
    If Not ActiveSheet Is wsS Then Exit Sub
    wsS.Cells(2, 2).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

' ================================================================
'  SCHNELLANSICHT AKTUALISIEREN
' ================================================================
Sub Schnellansicht_Aktualisieren()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsS As Worksheet: Set wsS = GetSheet("Schnell")
    If wsA Is Nothing Or wsS Is Nothing Then Exit Sub

    ' Spalten im Artikel-Blatt ueber die Ueberschrift finden - nie fest verdrahten,
    ' die Reihenfolge dort darf sich aendern.
    Dim colArt  As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    If colArt = 0 Then
        MsgBox "Spalte 'ARTIKEL' nicht gefunden! Bitte Spaltenüberschrift prüfen.", vbCritical
        Exit Sub
    End If
    Dim colEAN  As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colVK   As Long: colVK = Spalte_Finden(wsA, "VK-PREIS")
    Dim colAnz  As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim colEK   As Long: colEK = Spalte_Finden(wsA, "EK-PREIS")
    Dim colNr   As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colEinh As Long: colEinh = Spalte_Finden(wsA, "EINHEIT")
    Dim colLag  As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colWG   As Long: colWG = Spalte_Finden(wsA, "WARENGRUPPE")
    Dim colAttr As Long: colAttr = Spalte_Finden(wsA, "Attribut")

    Application.ScreenUpdating = False
    On Error GoTo Fehler

    ' --- 1. Altes weg: ALLE Spalten der Liste, nicht nur die vorderen ---------
    ' ⛔ Frueher wurde nur A:F geleert. Die Spalten G-M blieben stehen und
    '   gehoerten danach zur falschen Zeile. Deshalb hier bis SV_SPALTEN.
    Dim lastSvRow As Long, l2 As Long
    lastSvRow = wsS.Cells(wsS.Rows.count, SV_ARTIKEL).End(xlUp).Row
    l2 = wsS.Cells(wsS.Rows.count, SV_NR).End(xlUp).Row
    If l2 > lastSvRow Then lastSvRow = l2
    l2 = wsS.UsedRange.Row + wsS.UsedRange.Rows.count - 1
    If l2 > lastSvRow Then lastSvRow = l2
    If lastSvRow >= SV_DATEN_START Then
        wsS.Rows(SV_DATEN_START & ":" & lastSvRow).Hidden = False
        wsS.Range(wsS.Cells(SV_DATEN_START, 1), _
                  wsS.Cells(lastSvRow, SV_SPALTEN)).ClearContents
    End If
    If wsS.AutoFilterMode Then wsS.AutoFilterMode = False

    ' --- 2. Ueberschriften mitschreiben --------------------------------------
    ' Sie stehen bewusst hier und nicht nur im Setup: solange Kopf und Daten
    ' aus derselben Routine kommen, koennen sie nicht auseinanderlaufen.
    Schnellansicht_KopfSchreiben wsS

    ' --- 3. Artikel-Blatt in EINEM Zug einlesen ------------------------------
    Dim lastRow As Long: lastRow = LetzteZeile(wsA, colArt)
    If lastRow < ART_DATEN_START Then
        Application.ScreenUpdating = True
        Application.StatusBar = "Schnellansicht: keine Artikel gefunden."
        Exit Sub
    End If

    Dim maxSp As Long, sp As Variant, s As Long
    sp = Array(colEAN, colArt, colVK, colAnz, colEK, colNr, colEinh, colLag, colWG, colAttr)
    For s = 0 To UBound(sp)
        If sp(s) > maxSp Then maxSp = sp(s)
    Next s
    ' Mindestens zwei Spalten lesen: eine EINZELNE Zelle liefert keinen Bereich,
    ' sondern einen einzelnen Wert - dann liefe UBound(dat, 1) auf einen Fehler.
    If maxSp < 2 Then maxSp = 2

    Dim dat As Variant
    dat = wsA.Range(wsA.Cells(ART_DATEN_START, 1), wsA.Cells(lastRow, maxSp)).Value

    ' --- 4. Liste im Speicher bauen ------------------------------------------
    Dim n As Long: n = UBound(dat, 1)
    Dim aus() As Variant
    ReDim aus(1 To n, 1 To SV_SPALTEN)
    Dim z As Long, k As Long
    For z = 1 To n
        If Trim(CStr(dat(z, colArt))) <> "" Then
            k = k + 1
            aus(k, SV_NR) = k
            If colNr > 0 Then aus(k, SV_ARTNR) = dat(z, colNr)
            aus(k, SV_ARTIKEL) = dat(z, colArt)
            If colEAN > 0 Then aus(k, SV_EAN) = dat(z, colEAN)
            If colVK > 0 Then aus(k, SV_VK) = dat(z, colVK)
            If colEK > 0 Then aus(k, SV_EK) = dat(z, colEK)
            If colAnz > 0 Then aus(k, SV_BESTAND) = dat(z, colAnz)
            If colEinh > 0 Then aus(k, SV_EINHEIT) = dat(z, colEinh)
            If colLag > 0 Then aus(k, SV_LAGERORT) = dat(z, colLag)
            If colWG > 0 Then aus(k, SV_WARENGRUPPE) = dat(z, colWG)
            If colAttr > 0 Then aus(k, SV_ATTRIBUT) = dat(z, colAttr)
            ' Zeile im Artikel-Blatt: z zaehlt ab 1, die Daten beginnen bei ART_DATEN_START
            aus(k, SV_ZEILENVERWEIS) = z + ART_DATEN_START - 1
        End If
    Next z

    ' --- 5. In EINEM Zug schreiben -------------------------------------------
    ' Excel nimmt aus einem zu grossen Feld nur den Teil, der in den Bereich passt.
    If k > 0 Then
        wsS.Cells(SV_DATEN_START, 1).Resize(k, SV_SPALTEN).Value = aus
    End If
    wsS.Columns(SV_ZEILENVERWEIS).Hidden = True
    wsS.Cells(2, SV_TREFFER).Value = ""
    Schnellansicht_NachOben wsS

    Application.ScreenUpdating = True
    Application.StatusBar = "Schnellansicht: " & k & " Artikel aktualisiert."
    Exit Sub

Fehler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Fehler " & Err.Number & " beim Aufbauen der Schnellansicht:" & Chr(10) & _
           Err.Description, vbCritical, "Schnellansicht"
End Sub

' ================================================================
'  SCHNELLANSICHT - Ueberschriftenzeile schreiben
'  Nur von Schnellansicht_Aktualisieren gerufen. Steht extra, damit
'  Kopf und Spaltennummern nachweislich aus derselben Quelle kommen.
' ================================================================
Private Sub Schnellansicht_KopfSchreiben(wsS As Worksheet)
    With wsS
        .Cells(SV_KOPF_ZEILE, SV_NR).Value = "#"
        .Cells(SV_KOPF_ZEILE, SV_ARTNR).Value = "Art.-Nr."
        .Cells(SV_KOPF_ZEILE, SV_ARTIKEL).Value = "Artikel"
        .Cells(SV_KOPF_ZEILE, SV_EAN).Value = "EAN"
        .Cells(SV_KOPF_ZEILE, SV_VK).Value = "VK-Preis " & ChrW(8364)
        ' Der EK-Kopf wird von Schnellansicht_EK_Toggle geleert, wenn die Spalte
        ' ausgeblendet ist - dann hier nicht wieder hinschreiben.
        If .Cells(SV_DATEN_START, SV_EK).NumberFormat <> ";;;" Then
            .Cells(SV_KOPF_ZEILE, SV_EK).Value = "EK-Preis " & ChrW(8364)
        End If
        .Cells(SV_KOPF_ZEILE, SV_BESTAND).Value = "Bestand"
        .Cells(SV_KOPF_ZEILE, SV_EINHEIT).Value = "Einheit"
        .Cells(SV_KOPF_ZEILE, SV_LAGERORT).Value = "Lagerort"
        .Cells(SV_KOPF_ZEILE, SV_WARENGRUPPE).Value = "Warengruppe"
        .Cells(SV_KOPF_ZEILE, SV_ATTRIBUT).Value = "Attribut"
    End With
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
    Dim vkPreis As Double: vkPreis = val(Replace(vkStr, ",", "."))

    Dim ekStr As String
    ekStr = InputBox("EK-Preis (z.B. 5.00):", "Neuer Artikel 3/5", "0.00")
    Dim ekPreis As Double: ekPreis = val(Replace(ekStr, ",", "."))

    Dim mwstStr As String
    mwstStr = InputBox("MwSt % (Standard: 19):", "Neuer Artikel 4/5", "19")
    Dim mwst As Double: mwst = val(mwstStr)
    If mwst = 0 Then mwst = 19

    Dim anzStr As String
    anzStr = InputBox("Anfangsbestand:", "Neuer Artikel 4/5", "0")
    Dim anzahl As Double: anzahl = val(anzStr)

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
    nRow = LetzteZeile(wsA, colArt2) + 1

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
        bRow = wsB.Cells(wsB.Rows.count, 3).End(xlUp).Row + 1
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
    ' ⛔ Der Blattcode steht NUR in Schnellansicht_BlattcodeSchreiben. Frueher
    '   wurde er hier ein zweites Mal zusammengebaut - genau solche Doppelungen
    '   waren die Ursache des Spaltenbild-Fehlers vom 29.08.2026.
    If Not Schnellansicht_BlattcodeSchreiben() Then Exit Sub

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

    ' --- Artikel Toolbar (Shapes + Events) einrichten ---
    NeueModule.Setup_Artikel_Toolbar

    MsgBox "Setup fertig!" & Chr(10) & Chr(10) & _
           "Artikel-Sheet: Buttons in Zeile 3 aktiv (GITHUB, NEUER ARTIKEL, ZU-/ABGANG, ETIKETT, EK ausbl., FILTER LOESCHEN, SCHNELLANSICHT)" & Chr(10) & _
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
        Set wsI = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsI.Name = "Inventur"
    End If

    ' Dieses Setup raeumt das ganze Blatt ab. Wenn schon gezaehlt wurde, waere die
    ' Zaehlung weg - deshalb hier IMMER fragen, auch im stillen Setup-Durchlauf.
    Dim schonGezaehlt As Long: schonGezaehlt = 0
    Dim pruefBis As Long: pruefBis = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    Dim p As Long
    For p = INV_DATEN_START To pruefBis
        If Trim(CStr(wsI.Cells(p, 7).Value)) <> "" Then schonGezaehlt = schonGezaehlt + 1
    Next p
    If schonGezaehlt > 0 Then
        If MsgBox("ACHTUNG: In der Inventurliste stehen bereits " & schonGezaehlt & _
                  " gezaehlte Mengen." & Chr(10) & Chr(10) & _
                  "Dieses Setup baut das Blatt komplett neu auf - die Zaehlung geht dabei VERLOREN." & _
                  Chr(10) & Chr(10) & "Wirklich neu aufbauen?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, "Inventur-Setup") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False
    wsI.Cells.Clear
    wsI.Cells.Interior.ColorIndex = xlNone
    wsI.Cells.FormatConditions.Delete
    On Error Resume Next
    wsI.AutoFilterMode = False
    On Error GoTo 0

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
    ' K2:L2 - zeigt, was noch nicht gezaehlt wurde
    wsI.Range("K2:L2").Merge
    wsI.Cells(2, INV_COL_DATUM).Value = "NICHT GEZAEHLT"
    wsI.Cells(2, INV_COL_DATUM).Interior.Color = RGB(89, 89, 89)
    wsI.Cells(2, INV_COL_DATUM).Font.Color = RGB(255, 255, 255)
    wsI.Cells(2, INV_COL_DATUM).Font.Bold = True
    wsI.Cells(2, INV_COL_DATUM).HorizontalAlignment = xlCenter
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
    wsI.Cells(3, 5).Font.Bold = True
    wsI.Cells(3, 5).Font.Size = 11
    ' H3 = Soll-Anzeige
    wsI.Cells(3, 8).Interior.Color = hellgrau
    wsI.Cells(3, 8).Font.Bold = True
    wsI.Cells(3, 8).Font.Size = 12
    wsI.Cells(3, 8).HorizontalAlignment = xlCenter
    ' I3 = Gezaehlt-Eingabe. Deutlich abgesetzt, damit man auf einen Blick sieht,
    ' wo der Cursor steht - Excel zeigt bei einer markierten Zelle sonst nur einen
    ' duennen Rahmen, den man auf hellem Grund kaum erkennt.
    wsI.Cells(3, 9).Interior.Color = RGB(255, 255, 0)
    wsI.Cells(3, 9).Font.Bold = True
    wsI.Cells(3, 9).Font.Size = 14
    wsI.Cells(3, 9).HorizontalAlignment = xlCenter
    With wsI.Cells(3, 9).Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(180, 0, 0)
    End With
    ' J3 = EINTRAGEN Button
    wsI.Cells(3, 10).Value = "EINTRAGEN"
    wsI.Cells(3, 10).Interior.Color = orange
    wsI.Cells(3, 10).Font.Color = RGB(255, 255, 255)
    wsI.Cells(3, 10).Font.Bold = True
    wsI.Cells(3, 10).HorizontalAlignment = xlCenter

    ' K3 / L3 / K4:L4 = die Auswertungen und der Zaehl-Modus, damit man sie nicht
    ' ueber Alt+F8 heraussuchen muss
    With wsI.Cells(3, INV_COL_DATUM)
        .Value = "OFFENE LISTE"
        .Interior.Color = RGB(120, 120, 120)
        .Font.Color = RGB(255, 255, 255): .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With wsI.Cells(3, INV_COL_WER)
        .Value = "GEZAEHLTE"
        .Interior.Color = RGB(120, 120, 120)
        .Font.Color = RGB(255, 255, 255): .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    wsI.Range("K4:L4").Merge
    With wsI.Cells(4, INV_COL_DATUM)
        .Value = "ZAEHL-MODUS AN"
        .Interior.Color = RGB(150, 30, 30)
        .Font.Color = RGB(255, 255, 255): .Font.Bold = True: .Font.Size = 9
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ' Labels in Zeile 4
    ' A4:C4 = Fortschritt ("gezaehlt X von Y"), wird von Inventur_FortschrittAnzeigen gefuellt
    wsI.Range("A4:C4").Merge
    wsI.Cells(4, 1).Font.Size = 9
    wsI.Cells(4, 1).Font.Bold = True
    wsI.Cells(4, 1).Font.Color = RGB(60, 60, 60)
    wsI.Cells(4, 1).HorizontalAlignment = xlLeft
    wsI.Cells(4, 5).Value = "Artikel"
    wsI.Cells(4, 5).Font.Size = 8
    wsI.Cells(4, 5).Font.Color = RGB(120, 120, 120)
    wsI.Cells(4, 8).Value = "Soll"
    wsI.Cells(4, 8).Font.Size = 8
    wsI.Cells(4, 8).Font.Color = RGB(120, 120, 120)
    wsI.Cells(4, 9).Value = "Menge + ENTER"
    wsI.Cells(4, 9).Font.Size = 8
    wsI.Cells(4, 9).Font.Bold = True
    wsI.Cells(4, 9).Font.Color = RGB(180, 0, 0)
    wsI.Rows(3).RowHeight = 26
    wsI.Rows(4).RowHeight = 16

    ' --- Zeile 5: Spaltenkoepfe Liste ---
    ' K/L = Beleg (wann, von wem). M/N versteckt: M haelt die Artikelzeile, damit das
    ' Zurueckschreiben nicht ueber EAN oder Name suchen muss (39 Artikel haben keine EAN,
    ' 33 EANs und 273 Namen sind doppelt vergeben). N sichert den Bestand vor der Uebernahme.
    Dim hdr As Variant
    hdr = Array("Nr", "EAN", "Artikel", "Lagerort", "EK-Preis", "SOLL", "GEZAEHLT", _
                "DIFFERENZ", "EK-Wert", "Bemerkung", "Gezaehlt am", "Von", "ArtZeile", "Bestand vorher")
    Dim j As Integer
    For j = 0 To UBound(hdr)
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
    wsI.Columns(INV_COL_DATUM).ColumnWidth = 17
    wsI.Columns(INV_COL_WER).ColumnWidth = 10
    wsI.Columns(INV_COL_ZEILE).Hidden = True
    wsI.Columns(INV_COL_VORHER).Hidden = True

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
    c = c & "    End If" & Chr(10)
    ' Menge + ENTER traegt direkt ein - ohne diesen Zweig braucht jeder einzelne
    ' Artikel einen Mausklick auf EINTRAGEN.
    c = c & "    If Target.Address = ""$I$3"" Then" & Chr(10)
    c = c & "        On Error GoTo Fehler" & Chr(10)
    c = c & "        Application.EnableEvents = False" & Chr(10)
    c = c & "        LagerMakros.Inventur_MengeBestaetigt" & Chr(10)
    c = c & "        Application.EnableEvents = True" & Chr(10)
    c = c & "        Exit Sub" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "    Exit Sub" & Chr(10)
    c = c & "Fehler:" & Chr(10)
    c = c & "    Application.EnableEvents = True" & Chr(10)
    c = c & "End Sub" & Chr(10)
    ' Faerbt das Feld ein, in dem der Cursor gerade steht - sonst sieht man nicht,
    ' wo man tippt, und die Menge landet im falschen Feld.
    c = c & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10)
    c = c & "    On Error Resume Next" & Chr(10)
    c = c & "    LagerMakros.Inventur_FeldMarkieren Target" & Chr(10)
    c = c & "End Sub" & Chr(10)
    ' Cancel = True NUR auf den Knopf-Zellen setzen. Stand es davor pauschal am Anfang,
    ' war der Doppelklick im ganzen Blatt tot - man kam in keine Zelle mehr hinein,
    ' auch nicht ins Suchfeld.
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    If Target.Row = 2 And Target.Column >= 7 And Target.Column <= 8 Then" & Chr(10)
    c = c & "        Cancel = True: LagerMakros.Inventur_Befuellen" & Chr(10)
    ' Obergrenze noetig: ohne sie loest auch der Knopf in K2/L2 die Bestandsuebernahme aus.
    c = c & "    ElseIf Target.Row = 2 And Target.Column >= 9 And Target.Column <= 10 Then" & Chr(10)
    c = c & "        Cancel = True: LagerMakros.Inventur_BestaendeUebernehmen" & Chr(10)
    c = c & "    ElseIf Target.Row = 2 And Target.Column >= 11 And Target.Column <= 12 Then" & Chr(10)
    c = c & "        Cancel = True: LagerMakros.Inventur_NichtGezaehlt" & Chr(10)
    c = c & "    ElseIf Target.Row = 3 And Target.Column = 10 Then" & Chr(10)
    c = c & "        Cancel = True: LagerMakros.Inventur_Eintragen" & Chr(10)
    c = c & "    ElseIf Target.Row = 3 And Target.Column = 11 Then" & Chr(10)
    c = c & "        Cancel = True: LagerMakros.Inventur_OffeneListe" & Chr(10)
    c = c & "    ElseIf Target.Row = 3 And Target.Column = 12 Then" & Chr(10)
    c = c & "        Cancel = True: LagerMakros.Inventur_GezaehlteListe" & Chr(10)
    c = c & "    ElseIf Target.Row = 4 And Target.Column >= 11 And Target.Column <= 12 Then" & Chr(10)
    c = c & "        Cancel = True: LagerMakros.Inventur_Sperren" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c

    Application.ScreenUpdating = True
    If Not silent Then MsgBox "Inventur-Sheet erstellt!" & Chr(10) & Chr(10) & _
        "Zaehlen (nur Scanner und Zahlenblock, keine Maus):" & Chr(10) & _
        "  1. EAN in B3 scannen  ->  Artikel erscheint, Cursor springt auf die Menge" & Chr(10) & _
        "  2. Menge tippen + ENTER  ->  eingetragen, Cursor zurueck ins Suchfeld" & Chr(10) & _
        "  (bei mehreren Treffern kommt eine Auswahlliste)" & Chr(10) & Chr(10) & _
        "Knoepfe (Doppelklick):" & Chr(10) & _
        "- BEFUELLEN      = alle Artikel laden" & Chr(10) & _
        "- NICHT GEZAEHLT = zeigt nur die offenen Zeilen" & Chr(10) & _
        "- EINTRAGEN      = wie ENTER auf der Menge" & Chr(10) & _
        "- UEBERNEHMEN    = Bestaende aktualisieren (bucht mit)" & Chr(10) & Chr(10) & _
        "Der Fortschritt steht links ueber der Liste.", vbInformation
End Sub

' ================================================================
'  INVENTUR - ARTIKEL SUCHEN (via B3)
' ================================================================
Sub Inventur_Suchen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsI Is Nothing Then Exit Sub

    Dim such As String: such = Trim(wsI.Cells(3, 2).Value)
    ' Ergebnisfelder leeren. Das Mengenfeld MUSS mit geleert werden: bleibt dort ein alter
    ' Wert stehen und man tippt zufaellig denselben, aendert sich der Zellwert nicht -
    ' dann feuert kein Ereignis und der Artikel wird stillschweigend nicht eingetragen.
    wsI.Range("E3:G3").ClearContents
    wsI.Cells(3, 8).ClearContents
    wsI.Cells(3, 9).ClearContents
    ' Merker zuruecksetzen: solange nichts eindeutig gefunden ist, darf nichts eingetragen werden
    g_InvArtikelZeile = 0

    If such = "" Then Exit Sub

    ' Die eigentliche Suche steckt in Inventur_ArtikelFinden - das Zaehlblatt nutzt
    ' dieselbe Funktion, damit es nur EINE Suchlogik gibt.
    Dim meldung As String
    Dim gefZeile As Long: gefZeile = Inventur_ArtikelFinden(such, meldung)
    If gefZeile = 0 Then
        wsI.Cells(3, 5).Value = meldung
        Exit Sub
    End If

    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    wsI.Cells(3, 5).Value = wsA.Cells(gefZeile, colArt).Value
    If colAnz > 0 Then wsI.Cells(3, 8).Value = wsA.Cells(gefZeile, colAnz).Value
    g_InvArtikelZeile = gefZeile
    wsI.Cells(3, 9).Select
    ' ACHTUNG: KEIN Application.SendKeys "{F2}" hier - das schaltet in Excel den Nummernblock
    ' aus (bekannter Nebeneffekt, am 16.08. bei Frank aufgetreten). Beim Zaehlen werden
    ' die Mengen ueber den Zahlenblock eingetippt; ein abgeschaltetes NumLock waere
    ' schlimmer als ein fehlender blinkender Cursor.
    ' Sichtbar wird das aktive Feld ueber den Farbwechsel.
    Inventur_FeldMarkieren wsI.Cells(3, 9)
End Sub

' ================================================================
'  INVENTUR - MENGE MIT ENTER BESTAETIGT (Worksheet_Change auf I3)
' ================================================================
' Eigener Einstieg, damit das Leeren von I3 (Korrektur) nicht als Eingabe zaehlt.
Sub Inventur_MengeBestaetigt()
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsI Is Nothing Then Exit Sub
    If Trim(CStr(wsI.Cells(3, 9).Value)) = "" Then Exit Sub
    Inventur_Eintragen
End Sub

' ================================================================
'  INVENTUR - ARTIKEL SUCHEN (gemeinsam fuer Inventurliste und Zaehlblatt)
' ================================================================
' Rueckgabe: Zeilennummer im Artikel-Blatt. 0 = nichts eindeutig gefunden,
' dann steht in "meldung" der Grund zum Anzeigen.
' Artikelnummer/EAN als lesbaren Text. CStr allein reicht bei grossen Zahlen nicht -
' Excel zeigt sonst "9,00274E+12". Wird ueberall verwendet, wo eine Nummer angezeigt wird.
Function Inventur_NrText(wert As Variant) As String
    On Error Resume Next
    If IsNumeric(wert) And Trim(CStr(wert)) <> "" Then
        Inventur_NrText = Format(wert, "0")
    Else
        Inventur_NrText = Trim(CStr(wert))
    End If
    Err.Clear
End Function

' Sammelt die Treffer, ohne etwas anzuzeigen. Rueckgabe: Anzahl Treffer.
' tr() wird mit den Zeilennummern im Artikel-Blatt gefuellt (max. INV_MAX_TREFFER).
' Die Anzeige macht der Aufrufer - die Inventurliste per Auswahlfenster, das
' Zaehlblatt als Liste im Blatt.
Function Inventur_TrefferSammeln(such As String, ByRef tr() As Long, _
                                 ByRef meldung As String) As Long
    Inventur_TrefferSammeln = 0
    meldung = ""
    ReDim tr(1 To INV_MAX_TREFFER)
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then meldung = "-- Artikel-Blatt fehlt --": Exit Function
    such = Trim(such)
    If such = "" Then Exit Function

    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    If colArt = 0 Then meldung = "-- Spalte ARTIKEL fehlt --": Exit Function

    Dim lastA As Long: lastA = LetzteZeile(wsA, colArt)
    If lastA < ART_DATEN_START Then meldung = "-- Nicht gefunden --": Exit Function

    ' Such-Spalten in EINEM Zugriff lesen. Zelle fuer Zelle waeren das bei jedem
    ' Scan rund 14.000 Einzelzugriffe - beim Scannen deutlich spuerbar.
    Dim vonSp As Long: vonSp = colArt
    Dim bisSp As Long: bisSp = colArt
    If colEAN > 0 Then
        If colEAN < vonSp Then vonSp = colEAN
        If colEAN > bisSp Then bisSp = colEAN
    End If
    ' Eine Zeile mehr: sonst liefert .Value bei einer einzigen Zelle keinen Array
    Dim dat As Variant
    dat = wsA.Range(wsA.Cells(ART_DATEN_START, vonSp), wsA.Cells(lastA + 1, bisSp)).Value
    Dim iArt As Long: iArt = colArt - vonSp + 1
    Dim iEAN As Long: iEAN = 0
    If colEAN > 0 Then iEAN = colEAN - vonSp + 1

    Dim nurZahlen As Boolean: nurZahlen = (such = CStr(val(such)) And val(such) > 0)
    Dim suchL As String: suchL = LCase(such)

    ' Mehrwort-Suche mit UND-Logik - genau wie die gewohnte Suche im Artikel-Blatt:
    ' "bohrer 6" findet alles, wo BEIDE Woerter vorkommen, egal an welcher Stelle.
    Dim woerter() As String: woerter = Split(suchL, " ")

    Dim treffer As Long: treffer = 0
    Dim i As Long, w As Integer, passt As Boolean, suchIn As String
    For i = 1 To UBound(dat, 1)
        If nurZahlen And iEAN > 0 Then
            suchIn = LCase(CStr(dat(i, iArt)) & " " & CStr(dat(i, iEAN)))
        Else
            suchIn = LCase(CStr(dat(i, iArt)))
        End If
        If suchIn <> "" Then
            passt = True
            For w = 0 To UBound(woerter)
                If Trim(woerter(w)) <> "" Then
                    If InStr(suchIn, Trim(woerter(w))) = 0 Then passt = False: Exit For
                End If
            Next w
            If passt Then
                treffer = treffer + 1
                If treffer <= INV_MAX_TREFFER Then tr(treffer) = i + ART_DATEN_START - 1
            End If
        End If
    Next i

    If treffer = 0 Then meldung = "-- Nicht gefunden --"
    Inventur_TrefferSammeln = treffer
End Function

' Sucht und laesst bei mehreren Treffern per Auswahlfenster waehlen.
' Wird von der Inventurliste benutzt; das Zaehlblatt nimmt Inventur_TrefferSammeln.
Function Inventur_ArtikelFinden(such As String, ByRef meldung As String) As Long
    Inventur_ArtikelFinden = 0
    meldung = ""
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then meldung = "-- Artikel-Blatt fehlt --": Exit Function
    such = Trim(such)
    If such = "" Then Exit Function

    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colNr  As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")

    ' Suchen laesst Inventur_TrefferSammeln - damit gibt es die Logik nur einmal
    Dim tr() As Long
    Dim treffer As Long: treffer = Inventur_TrefferSammeln(such, tr, meldung)

    If treffer = 0 Then Exit Function
    If treffer = 1 Then Inventur_ArtikelFinden = tr(1): Exit Function
    If treffer > INV_MAX_TREFFER Then
        meldung = treffer & " Treffer - genauer suchen"
        Exit Function
    End If

    ' Mehrere Treffer: auswaehlen lassen statt in einer Sackgasse zu enden.
    ' 33 EANs und 273 Artikelnamen sind mehrfach vergeben - ohne Auswahl waeren
    ' diese Artikel gar nicht zaehlbar.
    Dim txt As String
    txt = treffer & " Treffer fuer """ & such & """:" & Chr(10) & Chr(10)
    Dim k As Long
    For k = 1 To treffer
        txt = txt & k & " = " & Left(CStr(wsA.Cells(tr(k), colArt).Value), 32)
        If colNr > 0 Then txt = txt & "  | Nr " & wsA.Cells(tr(k), colNr).Value
        If colLag > 0 Then txt = txt & "  | " & wsA.Cells(tr(k), colLag).Value
        If colAnz > 0 Then txt = txt & "  | Soll " & wsA.Cells(tr(k), colAnz).Value
        txt = txt & Chr(10)
    Next k
    txt = txt & Chr(10) & "Nummer eingeben (leer = abbrechen):"

    Dim wahl As String: wahl = InputBox(txt, "Welcher Artikel?")
    If Trim(wahl) = "" Then meldung = "-- abgebrochen --": Exit Function
    Dim wn As Long: wn = val(wahl)
    If wn < 1 Or wn > treffer Then meldung = "-- ungueltige Auswahl --": Exit Function
    Inventur_ArtikelFinden = tr(wn)
End Function

' ================================================================
'  INVENTUR - AKTIVES FELD SICHTBAR MACHEN
' ================================================================
' Excel zeichnet um die markierte Zelle nur einen duennen Rahmen, den man auf
' farbigem Grund kaum sieht - Frank hat deshalb am 16.08. die Menge ins falsche
' Feld getippt. SendKeys "{F2}" (Schreibmodus mit blinkendem Cursor) blieb bei ihm
' wirkungslos. Deshalb faerbt sich das aktive Feld jetzt selbst ein.
Sub Inventur_FeldMarkieren(Target As Range)
    On Error Resume Next
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsI Is Nothing Then Err.Clear: Exit Sub

    Dim aktivSuche As Boolean, aktivMenge As Boolean
    If Target.Row = 3 Then
        aktivSuche = (Target.Column >= 2 And Target.Column <= 4)
        aktivMenge = (Target.Column = 9)
    End If

    ' Nach ENTER schiebt Excel den Cursor eine Zeile nach unten - und zwar NACH dem
    ' Select im Code, das damit wirkungslos wird. Zeile 4 ist reine Beschriftung; wer
    ' dort landet, gehoert in eines der beiden Eingabefelder. Sonst landet der naechste
    ' Scan in der falschen Zelle.
    If Target.Row = 4 And Target.Column <= 9 Then
        Application.EnableEvents = False
        If g_InvArtikelZeile > 0 And Trim(CStr(wsI.Cells(3, 9).Value)) = "" Then
            wsI.Cells(3, 9).Select          ' Artikel gefunden -> Menge eingeben
            aktivSuche = False: aktivMenge = True
        Else
            wsI.Cells(3, 2).Select          ' sonst zurueck ins Suchfeld
            aktivSuche = True: aktivMenge = False
        End If
        Application.EnableEvents = True
    End If

    Dim kraeftig As Long: kraeftig = RGB(255, 255, 0)
    Dim blass    As Long: blass = RGB(240, 240, 240)

    ' Suchfeld B3:D3
    wsI.Range("B3:D3").Interior.Color = IIf(aktivSuche, kraeftig, blass)
    With wsI.Range("B3:D3").Borders
        .LineStyle = IIf(aktivSuche, xlContinuous, xlNone)
        .Weight = xlThick
        .Color = RGB(180, 0, 0)
    End With

    ' Mengenfeld I3
    wsI.Cells(3, 9).Interior.Color = IIf(aktivMenge, kraeftig, blass)
    With wsI.Cells(3, 9).Borders
        .LineStyle = IIf(aktivMenge, xlContinuous, xlNone)
        .Weight = xlThick
        .Color = RGB(180, 0, 0)
    End With

    ' Klartext daneben, damit man nicht raten muss, was als naechstes drankommt.
    ' NICHT nach B4 schreiben - A4:C4 ist die verbundene Fortschritts-Zelle. D4 ist frei.
    wsI.Cells(4, 4).Value = IIf(aktivSuche, "JETZT scannen", "")
    wsI.Cells(4, 4).Font.Bold = True
    wsI.Cells(4, 4).Font.Size = 8
    wsI.Cells(4, 4).Font.Color = RGB(180, 0, 0)
    wsI.Cells(4, 9).Value = IIf(aktivMenge, "JETZT Menge + ENTER", "Menge + ENTER")
    wsI.Cells(4, 9).Font.Bold = True
    wsI.Cells(4, 9).Font.Color = IIf(aktivMenge, RGB(180, 0, 0), RGB(150, 150, 150))
    Err.Clear
End Sub

' ================================================================
'  ZAEHLBLATT - abgeschottetes Blatt fuer den Mitarbeiter
' ================================================================
' Zweck: Wer zaehlt, soll NUR zaehlen koennen. Auf diesem Blatt sind genau zwei
' Zellen beschreibbar (Suchfeld und Menge); alles andere ist gesperrt, ebenso das
' Artikel-Blatt und die Inventurliste. Ein versehentliches Loeschen im Artikelstamm
' waere sonst nicht zu bemerken und wuerde 6.986 Artikel beschaedigen.
Sub Zaehlen_Setup(Optional silent As Boolean = False)
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsZ Is Nothing Then
        Set wsZ = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsZ.Name = "ZAEHLEN"
    End If

    Application.ScreenUpdating = False
    On Error Resume Next
    wsZ.Unprotect INV_SCHUTZ
    On Error GoTo 0
    wsZ.Cells.Clear
    wsZ.Cells.Interior.ColorIndex = xlNone

    Dim blau As Long:  blau = RGB(31, 56, 100)
    Dim grau As Long:  grau = RGB(242, 242, 242)
    Dim gelb As Long:  gelb = RGB(255, 255, 0)

    ' Spalten A und B bleiben als linker Rand frei, das Formular beginnt bei C.
    ' Zeile 1: Titel
    wsZ.Range("C1:I1").Merge
    With wsZ.Cells(1, ZB_LABEL)
        .Value = "INVENTUR  -  ZAEHLEN"
        .Interior.Color = blau: .Font.Color = RGB(255, 255, 255)
        .Font.Size = 18: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    wsZ.Rows(1).RowHeight = 42

    ' Zeile 2: Fortschritt
    wsZ.Range("C2:I2").Merge
    With wsZ.Cells(2, ZB_LABEL)
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(60, 60, 60)
        .HorizontalAlignment = xlCenter
    End With
    wsZ.Rows(2).RowHeight = 24

    ' Zeile 3: Hinweisbalken - sagt in Klartext, welches Feld gerade dran ist.
    ' Der blinkende Schreibcursor ist in Excel nicht erzwingbar (SendKeys "{F2}"
    ' schaltet den Nummernblock aus), deshalb dieser Balken.
    wsZ.Range("C3:I3").Merge
    With wsZ.Cells(3, ZB_LABEL)
        .Font.Size = 12: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    wsZ.Rows(3).RowHeight = 28

    ' Zeile 4/5: Suchfeld
    wsZ.Range("C4:I4").Merge
    wsZ.Cells(4, ZB_LABEL).Value = "1.  ARTIKEL SCANNEN   (oder Namen tippen)  und ENTER"
    wsZ.Cells(4, ZB_LABEL).Font.Bold = True
    wsZ.Cells(4, ZB_LABEL).Font.Size = 11
    wsZ.Cells(4, ZB_LABEL).Font.Color = blau
    wsZ.Rows(4).RowHeight = 20
    wsZ.Range(ZB_SUCHE).Merge
    With wsZ.Range(ZB_SUCHE)
        .NumberFormat = "@"
        .Interior.Color = gelb
        .Font.Size = 18: .Font.Bold = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    wsZ.Rows(5).RowHeight = 36
    wsZ.Rows(6).RowHeight = 14

    ' Zeile 7-9: Artikelanzeige - Felder sichtbar abgesetzt, sonst sieht man nicht,
    ' wo etwas steht und wo nicht
    wsZ.Cells(7, ZB_LABEL).Value = "Artikel"
    wsZ.Range(ZB_ARTIKEL).Merge
    With wsZ.Range(ZB_ARTIKEL)
        .Font.Size = 15: .Font.Bold = True: .Font.Color = blau
        .Interior.Color = grau
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    wsZ.Rows(7).RowHeight = 32

    wsZ.Cells(8, ZB_LABEL).Value = "Art.-Nr."
    wsZ.Cells(8, ZB_L2).Value = "EAN"
    wsZ.Cells(9, ZB_LABEL).Value = "Lagerort"
    wsZ.Range(ZB_LAGER).Merge
    wsZ.Range(ZB_EAN).Merge
    Dim feld As Variant
    feld = Array("D8", ZB_EAN, ZB_LAGER)
    Dim f As Integer
    For f = 0 To UBound(feld)
        With wsZ.Range(CStr(feld(f)))
            .Interior.Color = grau
            .Font.Size = 11
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(190, 190, 190)
        End With
    Next f
    wsZ.Rows(8).RowHeight = 22
    wsZ.Rows(9).RowHeight = 22
    wsZ.Rows(10).RowHeight = 16

    Dim r As Long
    For r = 7 To 9
        wsZ.Cells(r, ZB_LABEL).Font.Bold = True
        wsZ.Cells(r, ZB_LABEL).Font.Color = RGB(110, 110, 110)
        wsZ.Cells(r, ZB_LABEL).HorizontalAlignment = xlRight
        wsZ.Cells(r, ZB_LABEL).VerticalAlignment = xlCenter
    Next r
    wsZ.Cells(8, ZB_L2).Font.Bold = True
    wsZ.Cells(8, ZB_L2).Font.Color = RGB(110, 110, 110)
    wsZ.Cells(8, ZB_L2).HorizontalAlignment = xlRight
    wsZ.Cells(8, ZB_L2).VerticalAlignment = xlCenter

    ' Zeile 11: SOLL und MENGE - der Arbeitsbereich, deutlich groesser
    wsZ.Cells(11, ZB_LABEL).Value = "SOLL"
    wsZ.Cells(11, ZB_LABEL).Font.Bold = True: wsZ.Cells(11, ZB_LABEL).Font.Size = 12
    wsZ.Cells(11, ZB_LABEL).Font.Color = RGB(110, 110, 110)
    wsZ.Cells(11, ZB_LABEL).HorizontalAlignment = xlRight
    wsZ.Cells(11, ZB_LABEL).VerticalAlignment = xlCenter
    With wsZ.Cells(11, ZB_WERT)
        .Interior.Color = grau
        .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(190, 190, 190)
    End With
    wsZ.Cells(11, ZB_L2).Value = "2.  MENGE"
    wsZ.Cells(11, ZB_L2).Font.Bold = True: wsZ.Cells(11, ZB_L2).Font.Size = 12
    wsZ.Cells(11, ZB_L2).Font.Color = blau
    wsZ.Cells(11, ZB_L2).HorizontalAlignment = xlRight
    wsZ.Cells(11, ZB_L2).VerticalAlignment = xlCenter
    With wsZ.Cells(11, ZB_W2)
        .Interior.Color = gelb
        .Font.Size = 22: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    wsZ.Rows(11).RowHeight = 42

    ' Zeile 12: Differenz
    wsZ.Cells(12, ZB_L2).Value = "Differenz"
    wsZ.Cells(12, ZB_L2).Font.Bold = True
    wsZ.Cells(12, ZB_L2).Font.Color = RGB(110, 110, 110)
    wsZ.Cells(12, ZB_L2).HorizontalAlignment = xlRight
    wsZ.Cells(12, ZB_W2).Font.Bold = True
    wsZ.Cells(12, ZB_W2).Font.Size = 12
    wsZ.Cells(12, ZB_W2).HorizontalAlignment = xlCenter
    wsZ.Rows(12).RowHeight = 20
    wsZ.Rows(13).RowHeight = 18

    ' Zeile 14-19: Kontrollstreifen der letzten Eingaben, als abgesetzter Block
    wsZ.Range("C14:I14").Merge
    wsZ.Cells(14, ZB_LABEL).Value = "   zuletzt eingetragen"
    wsZ.Cells(14, ZB_LABEL).Font.Bold = True
    wsZ.Cells(14, ZB_LABEL).Font.Size = 9
    wsZ.Cells(14, ZB_LABEL).Font.Color = RGB(255, 255, 255)
    wsZ.Cells(14, ZB_LABEL).Interior.Color = RGB(150, 150, 150)
    wsZ.Rows(14).RowHeight = 16
    For r = 15 To 19
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, ZB_ENDE)).Merge
        With wsZ.Cells(r, ZB_LABEL)
            .Font.Size = 10
            .Font.Color = RGB(90, 90, 90)
            .Interior.Color = RGB(250, 250, 250)
            .IndentLevel = 1
            .VerticalAlignment = xlCenter
        End With
        wsZ.Rows(r).RowHeight = 17
    Next r
    With wsZ.Range("C14:I19").Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200)
    End With

    wsZ.Columns(1).ColumnWidth = 3      ' A - linker Rand
    wsZ.Columns(2).ColumnWidth = 4      ' B - Rand, traegt den Pfeil auf das aktive Feld
    wsZ.Columns(3).ColumnWidth = 11     ' C - Beschriftungen
    wsZ.Columns(4).ColumnWidth = 20     ' D - Werte
    wsZ.Columns(5).ColumnWidth = 18     ' E
    wsZ.Columns(6).ColumnWidth = 15     ' F - Beschriftung + Art.-Nr. in der Trefferliste
    wsZ.Columns(7).ColumnWidth = 16     ' G - Mengenfeld
    wsZ.Columns(8).ColumnWidth = 16     ' H
    wsZ.Columns(9).ColumnWidth = 16     ' I

    ' Nur die beiden Eingabefelder freigeben.
    ' ACHTUNG: IMMER den ganzen Merge-Bereich ansprechen (B5:F5). Auf einer Teilzelle wirft
    ' .Locked den Laufzeitfehler 1004 - dieselbe Falle wie bei ClearContents.
    wsZ.Cells.Locked = True
    wsZ.Range(ZB_SUCHE).Locked = False
    wsZ.Cells(11, ZB_W2).Locked = False

    ' --- Ereignisse ins Blatt schreiben ---
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsZ.CodeName)
    Dim cm As Object: Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim c As String: c = ""
    ' ACHTUNG: NICHT ueber Target.Address vergleichen: bei einer VERBUNDENEN Zelle meldet
    ' Excel den ganzen Bereich ("$D$5:$H$5"), nicht die Ankerzelle - der Vergleich
    ' schlaegt dann fehl und es passiert gar nichts. Intersect trifft immer.
    c = c & "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(10)
    c = c & "    If Not Intersect(Target, Me.Range(""D5"")) Is Nothing Then" & Chr(10)
    c = c & "        On Error GoTo Fehler" & Chr(10)
    c = c & "        Application.EnableEvents = False" & Chr(10)
    c = c & "        LagerMakros.Zaehlen_Suchen" & Chr(10)
    c = c & "        Application.EnableEvents = True" & Chr(10)
    c = c & "        Exit Sub" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "    If Not Intersect(Target, Me.Range(""G11"")) Is Nothing Then" & Chr(10)
    c = c & "        On Error GoTo Fehler" & Chr(10)
    c = c & "        Application.EnableEvents = False" & Chr(10)
    c = c & "        LagerMakros.Zaehlen_MengeBestaetigt" & Chr(10)
    c = c & "        Application.EnableEvents = True" & Chr(10)
    c = c & "        Exit Sub" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "    Exit Sub" & Chr(10)
    c = c & "Fehler:" & Chr(10)
    c = c & "    Application.EnableEvents = True" & Chr(10)
    c = c & "End Sub" & Chr(10)
    c = c & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & Chr(10)
    c = c & "    On Error Resume Next" & Chr(10)
    c = c & "    LagerMakros.Zaehlen_FeldMarkieren Target" & Chr(10)
    c = c & "End Sub" & Chr(10)
    ' Doppelklick = Excels Schreibmodus. Vorher MUSS der Platzhalter raus, sonst
    ' tippt man dahinter ("|  hier scannen ...Hundekissen") und die Suche findet nichts.
    c = c & "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & Chr(10)
    c = c & "    On Error Resume Next" & Chr(10)
    ' Doppelklick auf die Titelzeile = versteckter Ausstieg aus dem Zaehl-Modus.
    ' Bewusst kein sichtbarer Knopf - den wuerde der Mitarbeiter irgendwann druecken.
    c = c & "    If Target.Row = 1 Then" & Chr(10)
    c = c & "        Cancel = True" & Chr(10)
    c = c & "        LagerMakros.Zaehlen_ModusBeenden" & Chr(10)
    c = c & "        Exit Sub" & Chr(10)
    c = c & "    End If" & Chr(10)
    c = c & "    LagerMakros.Zaehlen_PlatzhalterWeg Target" & Chr(10)
    c = c & "End Sub" & Chr(10)
    cm.AddFromString c

    Zaehlen_Fortschritt
    wsZ.Protect Password:=INV_SCHUTZ, UserInterfaceOnly:=True
    wsZ.Activate
    wsZ.Range(ZB_SUCHE).Select
    ' Ohne Gitternetz und Ueberschriften sieht das Blatt wie ein Formular aus
    ' statt wie eine Tabelle - das ist der groesste optische Unterschied.
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = True

    If Not silent Then MsgBox "Zaehlblatt erstellt." & Chr(10) & Chr(10) & _
        "Auf diesem Blatt sind nur zwei Felder beschreibbar:" & Chr(10) & _
        "  - das gelbe Suchfeld" & Chr(10) & _
        "  - das gelbe Mengenfeld" & Chr(10) & Chr(10) & _
        "Alles andere ist gesperrt. Zum Aufheben:" & Chr(10) & _
        "  Strg+G  ->  Inventur_Entsperren", vbInformation
End Sub

' ---- Zaehlblatt: suchen -----------------------------------------
Sub Zaehlen_Suchen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsA Is Nothing Or wsZ Is Nothing Then Exit Sub

    ' ACHTUNG: Verbundene Zellen immer als ganzen Bereich ansprechen -
    ' auf einer Teilzelle bricht ClearContents mit Laufzeitfehler 1004 ab.
    Dim such As String: such = Trim(wsZ.Range(ZB_SUCHE).Cells(1, 1).Value)
    ' Der Platzhalter ist kein Suchbegriff - weder allein noch als Vorspann.
    ' Der Vorspann entsteht, wenn jemand doppelklickt und hinter den Platzhalter tippt.
    If such = ZB_PLATZ_SUCHE Then such = ""
    If Len(such) > Len(ZB_PLATZ_SUCHE) Then
        If Left(such, Len(ZB_PLATZ_SUCHE)) = ZB_PLATZ_SUCHE Then
            such = Trim(Mid(such, Len(ZB_PLATZ_SUCHE) + 1))
        End If
    End If
    wsZ.Range(ZB_ARTIKEL).ClearContents
    wsZ.Cells(8, ZB_WERT).ClearContents
    wsZ.Range(ZB_EAN).ClearContents
    wsZ.Range(ZB_LAGER).ClearContents
    wsZ.Cells(11, ZB_WERT).ClearContents
    wsZ.Cells(11, ZB_W2).ClearContents
    wsZ.Cells(12, ZB_W2).ClearContents
    g_InvArtikelZeile = 0
    If such = "" Then Exit Sub

    ' Treffer sammeln - die Anzeige macht dieses Blatt selbst, als Liste zum
    ' Anklicken. Ein Auswahlfenster mit Nummern ist bei 18 Treffern unbrauchbar.
    Dim meldung As String
    Dim tr() As Long
    Dim anzahl As Long: anzahl = Inventur_TrefferSammeln(such, tr, meldung)

    Zaehlen_ListeWeg

    If anzahl = 0 Then
        wsZ.Range(ZB_ARTIKEL).Value = meldung
        Exit Sub
    End If
    If anzahl > INV_MAX_TREFFER Then
        wsZ.Range(ZB_ARTIKEL).Value = anzahl & " Treffer - bitte genauer suchen"
        Exit Sub
    End If
    If anzahl > 1 Then
        Zaehlen_TrefferListe tr, anzahl
        Exit Sub
    End If

    Dim gefZeile As Long: gefZeile = tr(1)
    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr  As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")

    wsZ.Range(ZB_ARTIKEL).Value = wsA.Cells(gefZeile, colArt).Value
    If colNr > 0 Then
        wsZ.Cells(8, ZB_WERT).NumberFormat = "@"
        wsZ.Cells(8, ZB_WERT).Value = Inventur_NrText(wsA.Cells(gefZeile, colNr).Value)
    End If
    If colEAN > 0 Then
        wsZ.Range(ZB_EAN).NumberFormat = "0"
        wsZ.Range(ZB_EAN).Value = wsA.Cells(gefZeile, colEAN).Value
    End If
    If colLag > 0 Then wsZ.Range(ZB_LAGER).Value = wsA.Cells(gefZeile, colLag).Value
    If colAnz > 0 Then wsZ.Cells(11, ZB_WERT).Value = wsA.Cells(gefZeile, colAnz).Value

    g_InvArtikelZeile = gefZeile
    wsZ.Cells(11, ZB_W2).Select
    Zaehlen_FeldMarkieren wsZ.Cells(11, ZB_W2)
End Sub

' ---- Zaehlblatt: Trefferliste anzeigen --------------------------
' Zeigt die Treffer als anklickbare Liste - so wie die gewohnte Suche im
' Artikel-Blatt. Die Artikelzeile steht versteckt in Spalte J.
Sub Zaehlen_TrefferListe(tr() As Long, anzahl As Long)
    On Error Resume Next
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsA Is Nothing Or wsZ Is Nothing Then Err.Clear: Exit Sub

    Dim evAlt As Boolean: evAlt = Application.EnableEvents
    Application.EnableEvents = False
    Dim zGesch As Boolean: zGesch = Inventur_SchutzMerken(wsZ)

    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr  As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")

    wsZ.Cells(14, ZB_LABEL).Value = "   " & anzahl & " Treffer  -  bitte die Zeile anklicken"
    wsZ.Cells(14, ZB_LABEL).Interior.Color = RGB(180, 0, 0)

    Dim r As Long, k As Long
    For k = 1 To anzahl
        r = ZB_LISTE_START + k - 1
        wsZ.Rows(r).RowHeight = 19
        On Error Resume Next
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, ZB_ENDE)).UnMerge
        On Error GoTo 0
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, 5)).Merge
        wsZ.Cells(r, ZB_LABEL).Value = " " & wsA.Cells(tr(k), colArt).Value
        If colNr > 0 Then
            ' Textformat UND Format(...,"0"), sonst zeigt Excel lange Artikelnummern
            ' als "9,00274E+12" - CStr allein reicht bei grossen Zahlen nicht.
            wsZ.Cells(r, ZB_L2).NumberFormat = "@"
            wsZ.Cells(r, ZB_L2).Value = Inventur_NrText(wsA.Cells(tr(k), colNr).Value)
            wsZ.Cells(r, ZB_L2).HorizontalAlignment = xlLeft
        End If
        If colLag > 0 Then wsZ.Cells(r, ZB_W2).Value = wsA.Cells(tr(k), colLag).Value
        If colAnz > 0 Then wsZ.Cells(r, 8).Value = wsA.Cells(tr(k), colAnz).Value
        wsZ.Cells(r, 10).Value = tr(k)          ' versteckter Merker
        With wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, 8))
            .Font.Size = 11
            .Font.Bold = False
            .Font.Color = RGB(0, 0, 0)
            .Interior.Color = IIf(k Mod 2 = 0, RGB(242, 242, 242), RGB(255, 255, 255))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(200, 200, 200)
            .VerticalAlignment = xlCenter
        End With
        wsZ.Cells(r, 8).HorizontalAlignment = xlCenter
    Next k
    wsZ.Columns(10).Hidden = True

    ' Restliche Zeilen des Bereichs leeren
    For r = ZB_LISTE_START + anzahl To ZB_LISTE_START + INV_MAX_TREFFER - 1
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, 10)).ClearContents
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, 8)).Interior.ColorIndex = xlNone
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, 8)).Borders.LineStyle = xlNone
        wsZ.Rows(r).RowHeight = 17
    Next r

    wsZ.Range(ZB_ARTIKEL).Value = "bitte unten auswaehlen"
    Inventur_SchutzZurueck wsZ, zGesch
    Application.EnableEvents = evAlt
    Err.Clear
End Sub

' ---- Zaehlblatt: Treffer angeklickt ----------------------------
Sub Zaehlen_TrefferWaehlen(zeile As Long)
    On Error Resume Next
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsA Is Nothing Or wsZ Is Nothing Then Err.Clear: Exit Sub

    Dim az As Long: az = val(wsZ.Cells(zeile, 10).Value)
    If az < ART_DATEN_START Then Err.Clear: Exit Sub

    Dim evAlt As Boolean: evAlt = Application.EnableEvents
    Application.EnableEvents = False
    Dim zGesch As Boolean: zGesch = Inventur_SchutzMerken(wsZ)

    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colNr  As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")

    wsZ.Range(ZB_ARTIKEL).Value = wsA.Cells(az, colArt).Value
    If colNr > 0 Then
        wsZ.Cells(8, ZB_WERT).NumberFormat = "@"
        wsZ.Cells(8, ZB_WERT).Value = Inventur_NrText(wsA.Cells(az, colNr).Value)
    End If
    If colEAN > 0 Then
        wsZ.Range(ZB_EAN).NumberFormat = "0"
        wsZ.Range(ZB_EAN).Value = wsA.Cells(az, colEAN).Value
    End If
    If colLag > 0 Then wsZ.Range(ZB_LAGER).Value = wsA.Cells(az, colLag).Value
    If colAnz > 0 Then wsZ.Cells(11, ZB_WERT).Value = wsA.Cells(az, colAnz).Value

    g_InvArtikelZeile = az
    Zaehlen_ListeWeg
    Inventur_SchutzZurueck wsZ, zGesch
    Application.EnableEvents = evAlt

    wsZ.Cells(11, ZB_W2).Select
    Zaehlen_FeldMarkieren wsZ.Cells(11, ZB_W2)
    Err.Clear
End Sub

' ---- Zaehlblatt: Trefferliste weg, Kontrollstreifen zurueck -----
Sub Zaehlen_ListeWeg()
    On Error Resume Next
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsZ Is Nothing Then Err.Clear: Exit Sub
    If Trim(CStr(wsZ.Cells(14, ZB_LABEL).Value)) = "" Then Err.Clear: Exit Sub
    If InStr(wsZ.Cells(14, ZB_LABEL).Value, "Treffer") = 0 Then Err.Clear: Exit Sub

    Dim evAlt As Boolean: evAlt = Application.EnableEvents
    Application.EnableEvents = False
    Dim zGesch As Boolean: zGesch = Inventur_SchutzMerken(wsZ)

    Dim r As Long
    For r = ZB_LISTE_START To ZB_LISTE_START + INV_MAX_TREFFER - 1
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, 10)).ClearContents
        On Error Resume Next
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, ZB_ENDE)).UnMerge
        On Error GoTo 0
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, 10)).Borders.LineStyle = xlNone
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, 10)).Interior.ColorIndex = xlNone
    Next r

    ' Kontrollstreifen wiederherstellen
    wsZ.Cells(14, ZB_LABEL).Value = "   zuletzt eingetragen"
    wsZ.Cells(14, ZB_LABEL).Interior.Color = RGB(150, 150, 150)
    For r = ZB_LISTE_START To ZB_LISTE_START + 4
        wsZ.Range(wsZ.Cells(r, ZB_LABEL), wsZ.Cells(r, ZB_ENDE)).Merge
        With wsZ.Cells(r, ZB_LABEL)
            .Font.Size = 10
            .Font.Color = RGB(90, 90, 90)
            .Interior.Color = RGB(250, 250, 250)
            .IndentLevel = 1
            .VerticalAlignment = xlCenter
        End With
        wsZ.Rows(r).RowHeight = 17
    Next r
    With wsZ.Range(wsZ.Cells(14, ZB_LABEL), wsZ.Cells(ZB_LISTE_START + 4, ZB_ENDE)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200)
    End With

    Inventur_SchutzZurueck wsZ, zGesch
    Application.EnableEvents = evAlt
    Err.Clear
End Sub

' ---- Zaehlblatt: Zaehl-Modus beenden (Doppelklick auf den Titel) -----
' Versteckter Ausstieg mit Kennwort. Ein sichtbarer Knopf waere hier falsch:
' den wuerde der Mitarbeiter frueher oder spaeter druecken.
Sub Zaehlen_ModusBeenden()
    Dim eingabe As String
    eingabe = InputBox("Zaehl-Modus beenden und alle Blaetter wieder einblenden?" & Chr(10) & Chr(10) & _
                       "Kennwort:", "Inventur")
    If Trim(eingabe) = "" Then Exit Sub
    If eingabe <> INV_SCHUTZ Then
        MsgBox "Kennwort stimmt nicht.", vbExclamation, "Inventur"
        Exit Sub
    End If
    Inventur_Entsperren
End Sub

' ---- Zaehlblatt: Platzhalter vor dem Schreibmodus entfernen -----
' Wird aus Worksheet_BeforeDoubleClick gerufen. Ohne das tippt man beim Doppelklick
' hinter den Platzhalter und die Suche laeuft ins Leere - beim Scannen faellt es
' nicht auf, weil der Scanner das Feld komplett ueberschreibt.
Sub Zaehlen_PlatzhalterWeg(Target As Range)
    On Error Resume Next
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsZ Is Nothing Then Err.Clear: Exit Sub
    Dim evAlt As Boolean: evAlt = Application.EnableEvents
    Application.EnableEvents = False
    If Target.Row = 5 Then
        If Trim(CStr(wsZ.Range(ZB_SUCHE).Cells(1, 1).Value)) = ZB_PLATZ_SUCHE Then
            wsZ.Range(ZB_SUCHE).ClearContents
        End If
        wsZ.Range(ZB_SUCHE).Font.Color = RGB(0, 0, 0)
    ElseIf Target.Row = 11 And Target.Column = ZB_W2 Then
        If Trim(CStr(wsZ.Cells(11, ZB_W2).Value)) = ZB_PLATZ_MENGE Then
            wsZ.Cells(11, ZB_W2).ClearContents
        End If
        wsZ.Cells(11, ZB_W2).Font.Color = RGB(0, 0, 0)
    End If
    Application.EnableEvents = evAlt
    Err.Clear
End Sub

' ---- Zaehlblatt: Menge mit ENTER bestaetigt ---------------------
Sub Zaehlen_MengeBestaetigt()
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsZ Is Nothing Then Exit Sub
    Dim m As String: m = Trim(CStr(wsZ.Cells(11, ZB_W2).Value))
    ' Leeres Feld oder nur der Platzhalter: nichts eintragen
    If m = "" Or m = ZB_PLATZ_MENGE Then Exit Sub
    Zaehlen_Eintragen
End Sub

' ---- Zaehlblatt: eintragen (schreibt in die Inventurliste) ------
Sub Zaehlen_Eintragen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsZ Is Nothing Or wsI Is Nothing Then Exit Sub

    Dim artName As String: artName = Trim(CStr(wsZ.Range(ZB_ARTIKEL).Cells(1, 1).Value))
    Dim mengeStr As String: mengeStr = Trim(CStr(wsZ.Cells(11, ZB_W2).Value))
    If mengeStr = ZB_PLATZ_MENGE Then mengeStr = ""

    If g_InvArtikelZeile = 0 Then
        MsgBox "Bitte zuerst einen Artikel scannen.", vbExclamation
        wsZ.Cells(11, ZB_W2).ClearContents
        wsZ.Range(ZB_SUCHE).Select
        Exit Sub
    End If
    If mengeStr = "" Then Exit Sub
    Dim menge As Double: menge = val(mengeStr)
    Dim az As Long: az = g_InvArtikelZeile

    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If lastI < INV_DATEN_START Then
        MsgBox "Die Inventurliste ist LEER - es fehlt BEFUELLEN." & Chr(10) & _
               "Bitte Frank Bescheid geben, es wird nichts eingetragen.", vbCritical
        Exit Sub
    End If

    ' Zielzeile ueber den Zeilenverweis (Spalte M)
    Dim zielZeile As Long: zielZeile = 0
    Dim i As Long
    For i = INV_DATEN_START To lastI
        If val(wsI.Cells(i, INV_COL_ZEILE).Value) = az Then zielZeile = i: Exit For
    Next i
    If zielZeile = 0 Then
        MsgBox "Dieser Artikel steht nicht in der Inventurliste." & Chr(10) & _
               "Bitte Frank Bescheid geben.", vbExclamation
        wsZ.Cells(11, ZB_W2).ClearContents
        wsZ.Range(ZB_SUCHE).Select
        Exit Sub
    End If

    ' Schon gezaehlt? Dazuzaehlen oder ersetzen
    Dim altStr As String: altStr = Trim(CStr(wsI.Cells(zielZeile, 7).Value))
    Dim hinweis As String: hinweis = ""
    If altStr <> "" Then
        Dim alt As Double: alt = val(altStr)
        Dim antwort As Integer
        antwort = MsgBox(artName & Chr(10) & Chr(10) & _
            "Fuer diesen Artikel sind bereits " & Format(alt, "0") & " gezaehlt." & Chr(10) & Chr(10) & _
            "JA          = dazuzaehlen  (" & Format(alt, "0") & " + " & Format(menge, "0") & _
            " = " & Format(alt + menge, "0") & ")" & Chr(10) & _
            "NEIN        = ersetzen     (neu: " & Format(menge, "0") & ")" & Chr(10) & _
            "ABBRECHEN   = nichts aendern", _
            vbYesNoCancel + vbQuestion + vbDefaultButton1, "Schon gezaehlt")
        If antwort = vbCancel Then
            wsZ.Cells(11, ZB_W2).ClearContents
            wsZ.Range(ZB_SUCHE).Select
            Exit Sub
        End If
        If antwort = vbYes Then menge = alt + menge: hinweis = " (dazu)"
    End If

    ' Inventurliste ist gesperrt - kurz oeffnen, schreiben, wieder zu
    Dim wsIgeschuetzt As Boolean: wsIgeschuetzt = Inventur_SchutzMerken(wsI)
    wsI.Cells(zielZeile, 7).Value = menge
    Inventur_BelegSchreiben wsI, zielZeile
    Inventur_SchutzZurueck wsI, wsIgeschuetzt

    Dim soll As Double: soll = val(wsI.Cells(zielZeile, 6).Value)
    wsZ.Cells(12, ZB_W2).Value = Format(menge - soll, "+0;-0;0")

    Zaehlen_StreifenEintrag Format(Now, "HH:MM") & "  " & Left(artName, 38) & _
                            "   ->  " & Format(menge, "0") & hinweis

    ' Felder leeren fuer den naechsten Artikel (verbundene Bereiche ganz ansprechen)
    wsZ.Range(ZB_SUCHE).ClearContents
    wsZ.Range(ZB_ARTIKEL).ClearContents
    wsZ.Cells(8, ZB_WERT).ClearContents
    wsZ.Range("E8:F8").ClearContents
    wsZ.Range(ZB_LAGER).ClearContents
    wsZ.Cells(11, ZB_WERT).ClearContents
    wsZ.Cells(11, ZB_W2).ClearContents
    g_InvArtikelZeile = 0
    Zaehlen_Fortschritt
    wsZ.Range(ZB_SUCHE).Select
    Zaehlen_FeldMarkieren wsZ.Range(ZB_SUCHE)
End Sub

' ---- Zaehlblatt: Kontrollstreifen (letzte 5 Eingaben) -----------
Sub Zaehlen_StreifenEintrag(text As String)
    On Error Resume Next
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsZ Is Nothing Then Err.Clear: Exit Sub
    Dim r As Long
    For r = 19 To 16 Step -1
        wsZ.Cells(r, ZB_LABEL).Value = wsZ.Cells(r - 1, ZB_LABEL).Value
    Next r
    wsZ.Cells(15, ZB_LABEL).Value = text
    wsZ.Cells(15, ZB_LABEL).Font.Color = RGB(0, 0, 0)
    wsZ.Cells(15, ZB_LABEL).Font.Bold = True
    For r = 16 To 19
        wsZ.Cells(r, ZB_LABEL).Font.Color = RGB(140, 140, 140)
        wsZ.Cells(r, ZB_LABEL).Font.Bold = False
    Next r
    Err.Clear
End Sub

' ---- Zaehlblatt: Fortschritt ------------------------------------
Sub Zaehlen_Fortschritt()
    On Error Resume Next
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsZ Is Nothing Or wsI Is Nothing Then Err.Clear: Exit Sub

    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If lastI < INV_DATEN_START Then
        wsZ.Cells(2, ZB_LABEL).Value = "Inventurliste ist leer - bitte BEFUELLEN"
        Err.Clear: Exit Sub
    End If
    Dim dat As Variant
    dat = wsI.Range(wsI.Cells(INV_DATEN_START, 3), wsI.Cells(lastI, 7)).Value
    Dim gesamt As Long, gez As Long, i As Long
    For i = 1 To UBound(dat, 1)
        If Trim(CStr(dat(i, 1))) <> "" Then
            gesamt = gesamt + 1
            If Trim(CStr(dat(i, 5))) <> "" Then gez = gez + 1
        End If
    Next i
    wsZ.Cells(2, ZB_LABEL).Value = "gezaehlt " & Format(gez, "#,##0") & " von " & _
                            Format(gesamt, "#,##0") & "        noch offen: " & _
                            Format(gesamt - gez, "#,##0")
    Err.Clear
End Sub

' ---- Zaehlblatt: aktives Feld sichtbar machen -------------------
Sub Zaehlen_FeldMarkieren(Target As Range)
    On Error Resume Next
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsZ Is Nothing Then Err.Clear: Exit Sub

    ' Ereignisse aus, solange hier geschrieben wird: das Setzen der Platzhalter
    ' wuerde sonst Worksheet_Change ausloesen und eine Suche nach "| hier scannen"
    ' starten. Alter Zustand wird am Ende wiederhergestellt.
    Dim evAlt As Boolean: evAlt = Application.EnableEvents
    Application.EnableEvents = False

    ' ⛔ Titelzeile NICHT umleiten: dort sitzt der versteckte Ausstieg per Doppelklick.
    ' Ohne diese Ausnahme springt der Cursor beim ersten Klick zurueck ins Suchfeld,
    ' und der Doppelklick erreicht die Titelzeile nie.
    If Target.Row = 1 Then
        Application.EnableEvents = evAlt
        Err.Clear
        Exit Sub
    End If

    ' Klick in der Trefferliste? Dann Artikel waehlen - NICHT umleiten, sonst
    ' kaeme man an die Liste gar nicht heran.
    If Target.Row >= ZB_LISTE_START And Target.Row < ZB_LISTE_START + INV_MAX_TREFFER Then
        If val(wsZ.Cells(Target.Row, 10).Value) >= ART_DATEN_START Then
            Application.EnableEvents = evAlt
            Zaehlen_TrefferWaehlen Target.Row
            Err.Clear
            Exit Sub
        End If
    End If

    Dim aktivSuche As Boolean, aktivMenge As Boolean
    aktivSuche = (Target.Row = 5 And Target.Column >= ZB_WERT And Target.Column <= 8)
    aktivMenge = (Target.Row = 11 And Target.Column = ZB_W2)

    ' Wer irgendwo anders landet, wird ins passende Eingabefeld geholt -
    ' auf einem gesperrten Blatt kann man sonst nichts tun.
    If Not aktivSuche And Not aktivMenge Then
        Dim mengeJetzt As String: mengeJetzt = Trim(CStr(wsZ.Cells(11, ZB_W2).Value))
        If mengeJetzt = ZB_PLATZ_MENGE Then mengeJetzt = ""
        If g_InvArtikelZeile > 0 And mengeJetzt = "" Then
            wsZ.Cells(11, ZB_W2).Select: aktivMenge = True
        Else
            wsZ.Range(ZB_SUCHE).Select: aktivSuche = True
        End If
    End If

    ' Immer den ganzen Merge-Bereich faerben und rahmen - sonst haengt der rote
    ' Rahmen nur an der ersten Zelle und sieht aus wie ein Fehler.
    wsZ.Range(ZB_SUCHE).Interior.Color = IIf(aktivSuche, RGB(255, 255, 0), RGB(245, 245, 245))
    With wsZ.Range(ZB_SUCHE).Borders
        .LineStyle = xlContinuous
        .Weight = IIf(aktivSuche, xlThick, xlThin)
        .Color = IIf(aktivSuche, RGB(180, 0, 0), RGB(190, 190, 190))
    End With
    wsZ.Cells(11, ZB_W2).Interior.Color = IIf(aktivMenge, RGB(255, 255, 0), RGB(245, 245, 245))
    With wsZ.Cells(11, ZB_W2).Borders
        .LineStyle = xlContinuous
        .Weight = IIf(aktivMenge, xlThick, xlThin)
        .Color = IIf(aktivMenge, RGB(180, 0, 0), RGB(190, 190, 190))
    End With

    ' Beschriftungen mitziehen
    wsZ.Cells(4, ZB_LABEL).Value = IIf(aktivSuche, _
        "1.  >>> HIER SCANNEN <<<   (oder Namen tippen)  und ENTER", _
        "1.  Artikel scannen   (oder Namen tippen)  und ENTER")
    wsZ.Cells(11, ZB_L2).Value = IIf(aktivMenge, "2.  >>> MENGE", "2.  Menge")

    ' Der Hinweisbalken in Zeile 3 - das ist das Auffaelligste auf dem Blatt
    Dim balken As Range: Set balken = wsZ.Cells(3, ZB_LABEL)
    If aktivMenge Then
        balken.Value = ">>>   JETZT DIE MENGE EINTIPPEN  und ENTER   <<<"
        balken.Interior.Color = RGB(198, 239, 206)
        balken.Font.Color = RGB(0, 97, 0)
    ElseIf aktivSuche Then
        balken.Value = ">>>   JETZT DEN ARTIKEL SCANNEN   <<<"
        balken.Interior.Color = RGB(255, 242, 204)
        balken.Font.Color = RGB(150, 90, 0)
    Else
        balken.Value = ""
        balken.Interior.ColorIndex = xlNone
    End If

    ' --- Platzhalter im leeren aktiven Feld: sieht aus wie ein wartender Cursor ---
    Dim wSuche As String: wSuche = Trim(CStr(wsZ.Range(ZB_SUCHE).Cells(1, 1).Value))
    Dim wMenge As String: wMenge = Trim(CStr(wsZ.Cells(11, ZB_W2).Value))

    If aktivSuche Then
        If wSuche = "" Then
            wsZ.Range(ZB_SUCHE).Value = ZB_PLATZ_SUCHE
            wsZ.Range(ZB_SUCHE).Font.Color = RGB(170, 170, 170)
        Else
            If wSuche <> ZB_PLATZ_SUCHE Then wsZ.Range(ZB_SUCHE).Font.Color = RGB(0, 0, 0)
        End If
    Else
        ' Im inaktiven Feld darf kein Platzhalter stehenbleiben
        If wSuche = ZB_PLATZ_SUCHE Then wsZ.Range(ZB_SUCHE).ClearContents
        wsZ.Range(ZB_SUCHE).Font.Color = RGB(0, 0, 0)
    End If

    If aktivMenge Then
        If wMenge = "" Then
            wsZ.Cells(11, ZB_W2).Value = ZB_PLATZ_MENGE
            wsZ.Cells(11, ZB_W2).Font.Color = RGB(170, 170, 170)
        Else
            If wMenge <> ZB_PLATZ_MENGE Then wsZ.Cells(11, ZB_W2).Font.Color = RGB(0, 0, 0)
        End If
    Else
        If wMenge = ZB_PLATZ_MENGE Then wsZ.Cells(11, ZB_W2).ClearContents
        wsZ.Cells(11, ZB_W2).Font.Color = RGB(0, 0, 0)
    End If

    ' --- Pfeil in der freien Randspalte, der auf das aktive Feld zeigt ---
    wsZ.Cells(5, 2).ClearContents
    wsZ.Cells(11, 2).ClearContents
    Dim pfeilZeile As Long: pfeilZeile = 0
    If aktivSuche Then pfeilZeile = 5
    If aktivMenge Then pfeilZeile = 11
    If pfeilZeile > 0 Then
        With wsZ.Cells(pfeilZeile, 2)
            .Value = ChrW(9654)               ' >
            .Font.Size = 16
            .Font.Bold = True
            .Font.Color = RGB(180, 0, 0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If

    Application.EnableEvents = evAlt
    Err.Clear
End Sub

' ================================================================
'  OFFENE ARTIKEL - Liste nach Lagerort (Bildschirm + Ausdruck)
' ================================================================
' Bei 6.986 Artikeln und 34 Lagerorten ist die Frage nicht "was fehlt insgesamt",
' sondern "was fehlt HIER in diesem Regal". Diese Liste ist nach Lagerort sortiert
' und hat eine leere Spalte zum handschriftlichen Eintragen - falls der Scanner
' streikt oder man erst zaehlt und spaeter erfasst.
Sub Inventur_OffeneListe()
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsI Is Nothing Then MsgBox "Inventurliste nicht gefunden.", vbCritical: Exit Sub

    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If lastI < INV_DATEN_START Then MsgBox "Die Inventurliste ist leer.", vbInformation: Exit Sub

    Application.ScreenUpdating = False

    Dim wsO As Worksheet: Set wsO = GetSheet("Offen")
    If wsO Is Nothing Then
        Set wsO = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsO.Name = "Offen"
    End If
    Inventur_SchutzAus wsO
    On Error Resume Next
    wsO.AutoFilterMode = False
    On Error GoTo 0
    wsO.Cells.Clear
    wsO.Cells.Interior.ColorIndex = xlNone

    ' Daten in einem Zug lesen: A=Nr, B=EAN, C=Artikel, D=Lagerort, E=EK, F=SOLL, G=GEZAEHLT
    Dim dat As Variant
    dat = wsI.Range(wsI.Cells(INV_DATEN_START, 1), wsI.Cells(lastI, 7)).Value

    Dim blau As Long: blau = RGB(31, 56, 100)
    wsO.Range("A1:E1").Merge
    With wsO.Cells(1, 1)
        .Value = "NOCH NICHT GEZAEHLT   -   Stand " & Format(Now, "DD.MM.YYYY HH:MM")
        .Interior.Color = blau: .Font.Color = RGB(255, 255, 255)
        .Font.Size = 13: .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    wsO.Rows(1).RowHeight = 26

    Dim hdr As Variant: hdr = Array("Lagerort", "EAN", "Artikel", "Soll", "gezaehlt")
    Dim j As Integer
    For j = 0 To 4
        With wsO.Cells(3, j + 1)
            .Value = hdr(j)
            .Interior.Color = RGB(46, 80, 144): .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True: .HorizontalAlignment = xlCenter
        End With
    Next j

    Dim zeile As Long: zeile = 4
    Dim offen As Long, gesamt As Long
    Dim i As Long
    For i = 1 To UBound(dat, 1)
        If Trim(CStr(dat(i, 3))) <> "" Then
            gesamt = gesamt + 1
            If Trim(CStr(dat(i, 7))) = "" Then
                offen = offen + 1
                wsO.Cells(zeile, 1).Value = IIf(Trim(CStr(dat(i, 4))) = "", "(ohne Lagerort)", dat(i, 4))
                wsO.Cells(zeile, 2).NumberFormat = "0"
                wsO.Cells(zeile, 2).Value = dat(i, 2)
                wsO.Cells(zeile, 3).Value = dat(i, 3)
                wsO.Cells(zeile, 4).Value = dat(i, 6)
                zeile = zeile + 1
            End If
        End If
    Next i

    If offen = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Alle " & gesamt & " Artikel sind gezaehlt - nichts mehr offen.", vbInformation
        Exit Sub
    End If

    wsO.Cells(2, 1).Value = offen & " von " & gesamt & " Artikeln noch offen"
    wsO.Cells(2, 1).Font.Bold = True

    ' Nach Lagerort, dann Artikel sortieren
    wsO.Range(wsO.Cells(4, 1), wsO.Cells(zeile - 1, 5)).Sort _
        Key1:=wsO.Cells(4, 1), Order1:=xlAscending, _
        Key2:=wsO.Cells(4, 3), Order2:=xlAscending, Header:=xlNo

    ' Rahmen und Spaltenbreiten
    With wsO.Range(wsO.Cells(3, 1), wsO.Cells(zeile - 1, 5)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(160, 160, 160)
    End With
    wsO.Columns(1).ColumnWidth = 28
    wsO.Columns(2).ColumnWidth = 16
    wsO.Columns(3).ColumnWidth = 42
    wsO.Columns(4).ColumnWidth = 8
    wsO.Columns(5).ColumnWidth = 12
    ' Die Spalte "gezaehlt" bleibt leer - zum Eintragen mit dem Stift
    wsO.Range(wsO.Cells(4, 5), wsO.Cells(zeile - 1, 5)).Interior.Color = RGB(255, 255, 220)

    ' Druck: Querformat, Kopfzeile auf jeder Seite, auf Seitenbreite
    On Error Resume Next
    With wsO.PageSetup
        .Orientation = xlLandscape
        .PrintTitleRows = "$3:$3"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    Err.Clear
    On Error GoTo 0

    wsO.Activate
    wsO.Cells(4, 1).Select
    ActiveWindow.FreezePanes = False
    wsO.Rows(4).Select
    ActiveWindow.FreezePanes = True
    wsO.Cells(4, 1).Select
    Application.ScreenUpdating = True

    MsgBox offen & " Artikel sind noch nicht gezaehlt." & Chr(10) & Chr(10) & _
           "Die Liste steht im Blatt 'Offen', sortiert nach Lagerort." & Chr(10) & _
           "Die letzte Spalte ist frei zum Eintragen mit dem Stift." & Chr(10) & Chr(10) & _
           "Ausdrucken: Strg+P (Querformat ist schon eingestellt).", vbInformation
End Sub

' ================================================================
'  GEZAEHLTE ARTIKEL - Liste mit Beleg (Bildschirm + Ausdruck)
' ================================================================
' Gegenstueck zu Inventur_OffeneListe. Zeigt, was bereits gezaehlt wurde -
' mit Zeitstempel und Zaehler, neueste zuerst. Das ist die Kontrolle nach dem
' Zaehltag und zugleich die Grundlage fuer den Inventurbeleg.
Sub Inventur_GezaehlteListe()
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsI Is Nothing Then MsgBox "Inventurliste nicht gefunden.", vbCritical: Exit Sub

    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If lastI < INV_DATEN_START Then MsgBox "Die Inventurliste ist leer.", vbInformation: Exit Sub

    Application.ScreenUpdating = False

    Dim wsG As Worksheet: Set wsG = GetSheet("Gezaehlt")
    If wsG Is Nothing Then
        Set wsG = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsG.Name = "Gezaehlt"
    End If
    Inventur_SchutzAus wsG
    On Error Resume Next
    wsG.AutoFilterMode = False
    On Error GoTo 0
    wsG.Cells.Clear
    wsG.Cells.Interior.ColorIndex = xlNone

    ' A=Nr B=EAN C=Artikel D=Lagerort E=EK F=SOLL G=GEZAEHLT H=DIFF I=EK-Wert
    ' J=Bemerkung K=Gezaehlt am L=Von
    Dim dat As Variant
    dat = wsI.Range(wsI.Cells(INV_DATEN_START, 1), wsI.Cells(lastI, INV_COL_WER)).Value

    Dim blau As Long: blau = RGB(31, 56, 100)
    wsG.Range("A1:I1").Merge
    With wsG.Cells(1, 1)
        .Value = "GEZAEHLTE ARTIKEL   -   Stand " & Format(Now, "DD.MM.YYYY HH:MM")
        .Interior.Color = blau: .Font.Color = RGB(255, 255, 255)
        .Font.Size = 13: .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    wsG.Rows(1).RowHeight = 26

    Dim hdr As Variant
    hdr = Array("Gezaehlt am", "Von", "Lagerort", "EAN", "Artikel", _
                "Soll", "Gezaehlt", "Diff.", "EK-Wert")
    Dim j As Integer
    For j = 0 To UBound(hdr)
        With wsG.Cells(3, j + 1)
            .Value = hdr(j)
            .Interior.Color = RGB(46, 80, 144): .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True: .HorizontalAlignment = xlCenter
        End With
    Next j

    Dim zeile As Long: zeile = 4
    Dim gez As Long, gesamt As Long
    Dim summeDiff As Double, summeWert As Double
    Dim i As Long
    For i = 1 To UBound(dat, 1)
        If Trim(CStr(dat(i, 3))) <> "" Then
            gesamt = gesamt + 1
            If Trim(CStr(dat(i, 7))) <> "" Then
                gez = gez + 1
                wsG.Cells(zeile, 1).Value = dat(i, INV_COL_DATUM)
                wsG.Cells(zeile, 1).NumberFormat = "DD.MM.YYYY HH:MM"
                wsG.Cells(zeile, 2).Value = dat(i, INV_COL_WER)
                wsG.Cells(zeile, 3).Value = IIf(Trim(CStr(dat(i, 4))) = "", "(ohne Lagerort)", dat(i, 4))
                wsG.Cells(zeile, 4).NumberFormat = "0"
                wsG.Cells(zeile, 4).Value = dat(i, 2)
                wsG.Cells(zeile, 5).Value = dat(i, 3)
                wsG.Cells(zeile, 6).Value = dat(i, 6)
                wsG.Cells(zeile, 7).Value = dat(i, 7)
                Dim d As Double: d = val(dat(i, 7)) - val(dat(i, 6))
                wsG.Cells(zeile, 8).Value = d
                summeDiff = summeDiff + d
                Dim wert As Double: wert = val(dat(i, 5)) * val(dat(i, 7))
                wsG.Cells(zeile, 9).Value = wert
                wsG.Cells(zeile, 9).NumberFormat = "#,##0.00"
                summeWert = summeWert + wert
                ' Abweichungen faerben - die will man beim Durchsehen sofort finden
                If d < 0 Then
                    wsG.Cells(zeile, 8).Interior.Color = RGB(255, 199, 206)
                    wsG.Cells(zeile, 8).Font.Color = RGB(156, 0, 6)
                ElseIf d > 0 Then
                    wsG.Cells(zeile, 8).Interior.Color = RGB(255, 235, 156)
                    wsG.Cells(zeile, 8).Font.Color = RGB(156, 87, 0)
                End If
                wsG.Cells(zeile, 8).HorizontalAlignment = xlCenter
                zeile = zeile + 1
            End If
        End If
    Next i

    If gez = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Es ist noch nichts gezaehlt.", vbInformation
        Exit Sub
    End If

    wsG.Cells(2, 1).Value = gez & " von " & gesamt & " gezaehlt   -   " & _
                            (gesamt - gez) & " noch offen   -   " & _
                            "Summe Abweichung: " & Format(summeDiff, "+#,##0;-#,##0;0") & " Stueck"
    wsG.Cells(2, 1).Font.Bold = True

    ' Neueste zuerst
    wsG.Range(wsG.Cells(4, 1), wsG.Cells(zeile - 1, 9)).Sort _
        Key1:=wsG.Cells(4, 1), Order1:=xlDescending, Header:=xlNo

    ' Summenzeile
    Dim sumRow As Long: sumRow = zeile + 1
    wsG.Range(wsG.Cells(sumRow, 1), wsG.Cells(sumRow, 9)).Interior.Color = blau
    wsG.Range(wsG.Cells(sumRow, 1), wsG.Cells(sumRow, 9)).Font.Color = RGB(255, 255, 255)
    wsG.Range(wsG.Cells(sumRow, 1), wsG.Cells(sumRow, 9)).Font.Bold = True
    wsG.Range(wsG.Cells(sumRow, 1), wsG.Cells(sumRow, 7)).Merge
    wsG.Cells(sumRow, 1).Value = "SUMME der gezaehlten Ware (EK):"
    wsG.Cells(sumRow, 1).HorizontalAlignment = xlRight
    wsG.Cells(sumRow, 8).Value = summeDiff
    wsG.Cells(sumRow, 8).HorizontalAlignment = xlCenter
    wsG.Cells(sumRow, 9).Value = summeWert
    wsG.Cells(sumRow, 9).NumberFormat = "#,##0.00 " & ChrW(8364)
    wsG.Rows(sumRow).RowHeight = 22

    With wsG.Range(wsG.Cells(3, 1), wsG.Cells(zeile - 1, 9)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(160, 160, 160)
    End With
    wsG.Columns(1).ColumnWidth = 17
    wsG.Columns(2).ColumnWidth = 9
    wsG.Columns(3).ColumnWidth = 22
    wsG.Columns(4).ColumnWidth = 16
    wsG.Columns(5).ColumnWidth = 40
    wsG.Columns(6).ColumnWidth = 8
    wsG.Columns(7).ColumnWidth = 10
    wsG.Columns(8).ColumnWidth = 8
    wsG.Columns(9).ColumnWidth = 12

    On Error Resume Next
    With wsG.PageSetup
        .Orientation = xlLandscape
        .PrintTitleRows = "$3:$3"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    Err.Clear
    On Error GoTo 0

    wsG.Activate
    wsG.Cells(4, 1).Select
    ActiveWindow.FreezePanes = False
    wsG.Rows(4).Select
    ActiveWindow.FreezePanes = True
    wsG.Cells(4, 1).Select
    Application.ScreenUpdating = True

    MsgBox gez & " Artikel sind gezaehlt." & Chr(10) & Chr(10) & _
           "EK-Wert der gezaehlten Ware: " & Format(summeWert, "#,##0.00") & " " & ChrW(8364) & Chr(10) & _
           "Summe der Abweichungen: " & Format(summeDiff, "+#,##0;-#,##0;0") & " Stueck" & Chr(10) & Chr(10) & _
           "Die Liste steht im Blatt 'Gezaehlt', neueste zuerst." & Chr(10) & _
           "Abweichungen sind farbig hinterlegt." & Chr(10) & _
           "Ausdrucken: Strg+P (Querformat ist eingestellt).", vbInformation
End Sub

' ================================================================
'  BLATTSCHUTZ waehrend der Inventur
' ================================================================
Sub Inventur_SchutzAus(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect INV_SCHUTZ
    Err.Clear
End Sub

Sub Inventur_SchutzAn(ws As Worksheet)
    On Error Resume Next
    ' UserInterfaceOnly: Makros duerfen weiter schreiben, der Mensch nicht.
    ' Achtung: Diese Eigenschaft ueberlebt das Speichern NICHT - deshalb sperren
    ' und entsperren die Schreibroutinen jedes Mal selbst.
    ws.Protect Password:=INV_SCHUTZ, UserInterfaceOnly:=True
    Err.Clear
End Sub

' Merkt sich, ob ein Blatt geschuetzt war, und hebt den Schutz auf.
' ⛔ Gegenstueck ist Inventur_SchutzZurueck - NIE einfach Inventur_SchutzAn danach
' aufrufen: das schuetzt sonst ein Blatt, das vorher frei war. Genau so war die
' Inventurliste am 16.08. nach dem ersten Zaehlen dauerhaft gesperrt, und
' Inventur_Setup blieb bei Cells.Clear haengen.
Function Inventur_SchutzMerken(ws As Worksheet) As Boolean
    On Error Resume Next
    Inventur_SchutzMerken = ws.ProtectContents
    If Inventur_SchutzMerken Then ws.Unprotect INV_SCHUTZ
    Err.Clear
End Function

Sub Inventur_SchutzZurueck(ws As Worksheet, warGeschuetzt As Boolean)
    On Error Resume Next
    If warGeschuetzt Then ws.Protect Password:=INV_SCHUTZ, UserInterfaceOnly:=True
    Err.Clear
End Sub

' ---------------------------------------------------------------
'  Inventur-Modus AN: nur noch das Blatt ZAEHLEN ist erreichbar
' ---------------------------------------------------------------
' Bewusst ueber AUSBLENDEN statt ueber Blattschutz:
' ACHTUNG: Blattschutz mit UserInterfaceOnly ueberlebt das Speichern NICHT. Nach dem
'   naechsten Oeffnen koennten die Makros nicht mehr in die geschuetzten Blaetter
'   schreiben - BEFUELLEN und UEBERNEHMEN wuerden abbrechen.
' Unsichtbare Blaetter dagegen sind fuer Makros voll beschreibbar, fuer den Menschen
' aber nicht erreichbar. xlSheetVeryHidden laesst sich auch nicht ueber Rechtsklick
' auf die Blattleiste einblenden, und der Mappenschutz verhindert das Einfuegen.
Sub Inventur_Sperren()
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")
    If wsZ Is Nothing Then
        MsgBox "Das Blatt ZAEHLEN fehlt." & Chr(10) & _
               "Bitte zuerst  Strg+G  ->  Zaehlen_Setup  ausfuehren.", vbExclamation
        Exit Sub
    End If
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsI Is Nothing Then MsgBox "Inventurliste fehlt.", vbCritical: Exit Sub
    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If lastI < INV_DATEN_START Then
        MsgBox "Die Inventurliste ist noch LEER." & Chr(10) & Chr(10) & _
               "Bitte zuerst BEFUELLEN druecken - sonst kann nicht gezaehlt werden.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Unprotect INV_SCHUTZ
    Err.Clear
    On Error GoTo 0

    ' Merken, welche Blaetter jetzt sichtbar sind - beim Entsperren wird genau
    ' dieser Zustand wiederhergestellt (ArtikelDetail & Co. sind normal versteckt).
    Dim merk As String: merk = ""
    Dim ws As Worksheet, n As Integer
    For Each ws In ThisWorkbook.Sheets
        If ws.Visible = xlSheetVisible Then merk = merk & ws.Name & "|"
    Next ws
    Dim zGesch As Boolean: zGesch = Inventur_SchutzMerken(wsZ)
    wsZ.Cells(30, ZB_LABEL).Value = merk
    wsZ.Rows(30).Hidden = True

    ' Alles ausser ZAEHLEN unsichtbar
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> wsZ.Name Then
            On Error Resume Next
            ws.Visible = xlSheetVeryHidden
            If Err.Number = 0 Then n = n + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next ws

    ' Zaehlblatt: nur die beiden Eingabefelder frei (ganzer Merge-Bereich, siehe oben)
    wsZ.Cells.Locked = True
    wsZ.Range(ZB_SUCHE).Locked = False
    wsZ.Cells(11, ZB_W2).Locked = False
    wsZ.Protect Password:=INV_SCHUTZ, UserInterfaceOnly:=True

    ' Mappenschutz: keine Blaetter einfuegen, verschieben oder einblenden
    On Error Resume Next
    ThisWorkbook.Protect Password:=INV_SCHUTZ, Structure:=True, Windows:=False
    Err.Clear
    On Error GoTo 0

    wsZ.Activate
    wsZ.Range(ZB_SUCHE).Select
    Zaehlen_FeldMarkieren wsZ.Range(ZB_SUCHE)

    MsgBox "Inventur-Modus AN." & Chr(10) & Chr(10) & _
           n & " Blaetter sind ausgeblendet - es gibt nur noch ZAEHLEN." & Chr(10) & _
           "Dort sind genau zwei Felder beschreibbar: Suche und Menge." & Chr(10) & Chr(10) & _
           "Aufheben:  Strg+G  ->  Inventur_Entsperren", vbInformation
End Sub

' ---------------------------------------------------------------
'  Inventur-Modus AUS: alles wieder wie vorher
' ---------------------------------------------------------------
Sub Inventur_Entsperren()
    Dim wsZ As Worksheet: Set wsZ = GetSheet("ZAEHLEN")

    On Error Resume Next
    ThisWorkbook.Unprotect INV_SCHUTZ
    Err.Clear
    On Error GoTo 0

    ' Gemerkten Sichtbarkeits-Zustand wiederherstellen
    Dim merk As String: merk = ""
    If Not wsZ Is Nothing Then
        Dim zGesch As Boolean: zGesch = Inventur_SchutzMerken(wsZ)
        wsZ.Rows(30).Hidden = False
        merk = CStr(wsZ.Cells(30, ZB_LABEL).Value)
        wsZ.Cells.Locked = False
    End If

    Dim ws As Worksheet, n As Integer
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        If merk = "" Then
            ' Nichts gemerkt (z.B. nie gesperrt): alles sichtbar machen ausser den
            ' Popup-Blaettern, die von Haus aus versteckt sind.
            If InStr(ws.Name, "Detail") = 0 And InStr(ws.Name, "Popup") = 0 _
               And ws.Name <> "Neuer Artikel" Then ws.Visible = xlSheetVisible
        ElseIf InStr(merk, ws.Name & "|") > 0 Then
            ws.Visible = xlSheetVisible
        End If
        If Err.Number = 0 Then n = n + 1
        Err.Clear
        On Error GoTo 0
    Next ws

    MsgBox "Inventur-Modus AUS - die Blaetter sind wieder da." & Chr(10) & Chr(10) & _
           "Die Zaehlung in der Inventurliste ist unveraendert.", vbInformation
End Sub

' ================================================================
'  INVENTUR - FORTSCHRITT ANZEIGEN (A4 und Statusleiste)
' ================================================================
' Nur Anzeige - Fehler hier duerfen den Aufrufer nicht stoeren. Deshalb Resume Next,
' und am Ende IMMER Err.Clear: ein stehengebliebener Err.Number laesst den Aufrufer
' sonst faelschlich einen Abbruch melden.
Sub Inventur_FortschrittAnzeigen(wsI As Worksheet)
    On Error Resume Next
    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If lastI < INV_DATEN_START Then
        wsI.Cells(4, 1).Value = ""
        Err.Clear
        Exit Sub
    End If

    Dim dat As Variant
    dat = wsI.Range(wsI.Cells(INV_DATEN_START, 3), wsI.Cells(lastI, 7)).Value
    Dim gesamt As Long, gez As Long
    Dim i As Long
    For i = 1 To UBound(dat, 1)
        If Trim(CStr(dat(i, 1))) <> "" Then
            gesamt = gesamt + 1
            If Trim(CStr(dat(i, 5))) <> "" Then gez = gez + 1
        End If
    Next i

    wsI.Cells(4, 1).Value = "gezaehlt " & Format(gez, "#,##0") & " von " & _
                            Format(gesamt, "#,##0") & "   (offen " & _
                            Format(gesamt - gez, "#,##0") & ")"
    Err.Clear
End Sub

' ================================================================
'  INVENTUR - MENGE EINTRAGEN
' ================================================================
Sub Inventur_Eintragen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsI Is Nothing Then Exit Sub

    Dim artName As String: artName = Trim(wsI.Cells(3, 5).Value)
    Dim mengeStr As String: mengeStr = Trim(wsI.Cells(3, 9).Value)

    ' Zugeordnet wird ueber die gemerkte Artikelzeile, nicht ueber den Namen:
    ' 273 Artikelnamen kommen mehrfach vor (700 Zeilen).
    If g_InvArtikelZeile = 0 Then
        MsgBox "Bitte zuerst einen Artikel suchen." & Chr(10) & Chr(10) & _
               "Wenn oben mehrere Treffer stehen, ist die Zuordnung nicht eindeutig -" & Chr(10) & _
               "dann bitte genauer suchen.", vbExclamation
        Exit Sub
    End If
    If mengeStr = "" Then
        MsgBox "Bitte Menge eingeben.", vbExclamation: Exit Sub
    End If
    Dim menge As Double: menge = val(mengeStr)

    Dim az As Long: az = g_InvArtikelZeile

    ' Artikel in der Liste ueber den Zeilenverweis (Spalte M) suchen
    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row

    ' Leere Liste = BEFUELLEN wurde vergessen (passiert leicht nach Inventur_Setup).
    ' Ohne Warnung baut sich die Liste still von unten neu auf und man zaehlt gegen
    ' eine Liste, in der die anderen 6.985 Artikel gar nicht stehen.
    If lastI < INV_DATEN_START Then
        If MsgBox("Die Inventurliste ist noch LEER - vermutlich fehlt BEFUELLEN." & Chr(10) & Chr(10) & _
                  "Wenn du jetzt eintraegst, entsteht eine Liste mit nur diesem einen Artikel." & _
                  Chr(10) & Chr(10) & "Trotzdem eintragen?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, "Liste ist leer") = vbNo Then Exit Sub
    End If
    Dim zielZeile As Long: zielZeile = 0
    Dim i As Long
    For i = INV_DATEN_START To lastI
        If val(wsI.Cells(i, INV_COL_ZEILE).Value) = az Then zielZeile = i: Exit For
    Next i

    ' Steht dort schon eine Zaehlung? Dann nicht stillschweigend ueberschreiben -
    ' derselbe Artikel kann an zwei Lagerorten liegen.
    If zielZeile > 0 Then
        Dim altStr As String: altStr = Trim(CStr(wsI.Cells(zielZeile, 7).Value))
        If altStr <> "" Then
            Dim alt As Double: alt = val(altStr)
            Dim antwort As Integer
            antwort = MsgBox( _
                artName & Chr(10) & Chr(10) & _
                "Fuer diesen Artikel sind bereits " & Format(alt, "0") & " gezaehlt." & Chr(10) & Chr(10) & _
                "JA          = dazuzaehlen  (" & Format(alt, "0") & " + " & Format(menge, "0") & _
                " = " & Format(alt + menge, "0") & ")" & Chr(10) & _
                "NEIN        = ersetzen     (neu: " & Format(menge, "0") & ")" & Chr(10) & _
                "ABBRECHEN   = nichts aendern", _
                vbYesNoCancel + vbQuestion + vbDefaultButton1, "Schon gezaehlt")
            If antwort = vbCancel Then Exit Sub
            If antwort = vbYes Then menge = alt + menge
        End If
    End If

    If zielZeile > 0 Then
        wsI.Cells(zielZeile, 7).Value = menge
        Inventur_BelegSchreiben wsI, zielZeile
        ' Kein .Select auf die Listenzeile: bei gesetztem Filter ist sie ausgeblendet
        ' und Select bricht mit einem Laufzeitfehler ab. Rueckmeldung unten in der Leiste.
        Application.StatusBar = "Gezaehlt: " & artName & "  ->  " & Format(menge, "0")
    Else
        ' Nicht in der Liste (z.B. nach dem Befuellen neu angelegt): unten anhaengen
        Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
        Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
        Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
        Dim colEK  As Long: colEK = Spalte_Finden(wsA, "EK-PREIS")
        Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
        Dim nRow   As Long
        If lastI < INV_DATEN_START Then nRow = INV_DATEN_START Else nRow = lastI + 1

        ' Nicht in die Summenzeile hineinschreiben - davor eine Zeile einfuegen
        Dim sumRow As Long: sumRow = Inventur_SummenZeile(wsI)
        If sumRow > 0 And nRow >= sumRow Then
            wsI.Rows(sumRow).Insert Shift:=xlDown
            nRow = sumRow
            sumRow = sumRow + 1
        End If

        wsI.Cells(nRow, 1).Value = nRow - INV_DATEN_START + 1
        If colEAN > 0 Then
            wsI.Cells(nRow, 2).NumberFormat = "0"   ' sonst steht dort "5,90214E+12"
            wsI.Cells(nRow, 2).Value = wsA.Cells(az, colEAN).Value
        End If
        wsI.Cells(nRow, 3).Value = wsA.Cells(az, colArt).Value
        If colLag > 0 Then wsI.Cells(nRow, 4).Value = wsA.Cells(az, colLag).Value
        If colEK > 0 Then wsI.Cells(nRow, 5).Value = wsA.Cells(az, colEK).Value
        If colAnz > 0 Then wsI.Cells(nRow, 6).Value = wsA.Cells(az, colAnz).Value
        wsI.Cells(nRow, 7).Value = menge
        wsI.Cells(nRow, 8).Formula = "=IF(G" & nRow & "="""","""",G" & nRow & "-F" & nRow & ")"
        wsI.Cells(nRow, 9).Formula = "=IF(G" & nRow & "<>"""",E" & nRow & "*G" & nRow & ",E" & nRow & "*F" & nRow & ")"
        wsI.Cells(nRow, 9).NumberFormat = "0.00"
        wsI.Cells(nRow, INV_COL_ZEILE).Value = az
        Inventur_BelegSchreiben wsI, nRow

        ' Summenformel nachziehen, damit die neue Zeile mitzaehlt
        If sumRow > 0 Then
            wsI.Cells(sumRow, 9).Formula = "=SUM(I" & INV_DATEN_START & ":I" & (sumRow - 1) & ")"
        End If
        Application.StatusBar = "Neu aufgenommen: " & artName & "  ->  " & Format(menge, "0")
    End If

    ' Suchfelder leeren fuer naechsten Artikel.
    ' IMMER den ganzen Merge-Bereich ansprechen (B3:D3, E3:G3), nicht die Teilzelle -
    ' ClearContents auf einem Teil einer verbundenen Zelle bricht ab. Der Blatt-Handler
    ' faengt jeden Fehler ab, deshalb blieb das frueher unbemerkt: die Zeile war
    ' geschrieben, aber Suchfeld und Fortschritt standen still.
    ' Fortschritt noch INNERHALB des Events-Aus schreiben - sonst loest das Schreiben
    ' in A4 gleich wieder den Blatt-Handler aus.
    Application.EnableEvents = False
    On Error Resume Next
    wsI.Range("B3:D3").ClearContents
    wsI.Range("E3:G3").ClearContents
    wsI.Cells(3, 8).ClearContents
    wsI.Cells(3, 9).ClearContents
    If Err.Number <> 0 Then
        Dim leerFehler As String: leerFehler = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Inventur_FortschrittAnzeigen wsI
    Application.EnableEvents = True
    g_InvArtikelZeile = 0
    On Error Resume Next
    wsI.Cells(3, 2).Select
    Inventur_FeldMarkieren wsI.Cells(3, 2)
    Err.Clear
    On Error GoTo 0
    If leerFehler <> "" Then
        MsgBox "Der Artikel wurde eingetragen, aber die Suchfelder liessen sich nicht " & _
               "leeren:" & Chr(10) & leerFehler & Chr(10) & Chr(10) & _
               "Bitte Frank Bescheid geben.", vbExclamation
    End If
End Sub

' ================================================================
'  INVENTUR - HILFSFUNKTIONEN
' ================================================================
' Schreibt Zeitstempel und Zaehler an eine Listenzeile (der Inventurbeleg).
Sub Inventur_BelegSchreiben(wsI As Worksheet, zeile As Long)
    wsI.Cells(zeile, INV_COL_DATUM).Value = Now()
    wsI.Cells(zeile, INV_COL_DATUM).NumberFormat = "DD.MM.YYYY HH:MM"
    wsI.Cells(zeile, INV_COL_WER).Value = BENUTZER
End Sub

' Findet die Summenzeile am Listenende (0 = keine vorhanden).
Function Inventur_SummenZeile(wsI As Worksheet) As Long
    Dim letzte As Long: letzte = wsI.Cells(wsI.Rows.count, 1).End(xlUp).Row
    Dim i As Long
    For i = letzte To INV_DATEN_START Step -1
        If InStr(1, CStr(wsI.Cells(i, 1).Value), "SUMME", vbTextCompare) > 0 Then
            Inventur_SummenZeile = i
            Exit Function
        End If
    Next i
    Inventur_SummenZeile = 0
End Function

' ================================================================
'  INVENTUR - DATEN BEFUELLEN
' ================================================================
Sub Inventur_Befuellen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsA Is Nothing Or wsI Is Nothing Then Exit Sub

    If MsgBox("Vorhandene Inventurdaten loeschen und neu laden?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    ' Ohne diese drei Schalter feuert bei rund 49.000 geschriebenen Zellen jedes Mal
    ' der Blatt-Handler und Excel rechnet alle Formeln neu - das dauert Minuten.
    Dim altCalcB As Long: altCalcB = Application.Calculation
    On Error GoTo BefuellenAufraeumen
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Alten Inhalt restlos raeumen - inklusive Summenzeile und der Spalten K-N.
    ' Ueber Spalte 3 allein wuerde die Summenzeile stehen bleiben (dort gemergt und leer)
    ' und beim naechsten Lauf als Geisterzeile mitten in der Liste landen.
    Dim letzteA As Long: letzteA = wsI.Cells(wsI.Rows.count, 1).End(xlUp).Row
    Dim letzteC As Long: letzteC = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If letzteC > letzteA Then letzteA = letzteC
    If letzteA >= INV_DATEN_START Then
        With wsI.Range(wsI.Cells(INV_DATEN_START, 1), wsI.Cells(letzteA, INV_COL_LETZTE))
            .UnMerge
            .Clear
            .FormatConditions.Delete
        End With
    End If

    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colEK  As Long: colEK = Spalte_Finden(wsA, "EK-PREIS")
    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    If colArt = 0 Then
        Application.EnableEvents = True
        Application.Calculation = altCalcB
        Application.ScreenUpdating = True
        MsgBox "Spalte 'ARTIKEL' im Artikel-Blatt nicht gefunden.", vbCritical
        Exit Sub
    End If

    Dim lastA As Long: lastA = LetzteZeile(wsA, colArt)
    Dim sRow As Long: sRow = INV_DATEN_START
    Dim nr As Long: nr = 1
    Dim hellgrau As Long: hellgrau = RGB(242, 242, 242)

    Dim i As Long
    For i = ART_DATEN_START To lastA
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
            ' Merker: aus welcher Artikelzeile stammt diese Listenzeile
            wsI.Cells(sRow, INV_COL_ZEILE).Value = i
            If sRow Mod 2 = 0 Then wsI.Range("A" & sRow & ":L" & sRow).Interior.Color = hellgrau
            sRow = sRow + 1
            nr = nr + 1
        End If
    Next i

    Dim letzteDaten As Long: letzteDaten = sRow - 1

    ' EAN als ganze Zahl zeigen. Ohne das steht in der Liste "5,90017E+12" und man
    ' erkennt den gescannten Artikel nicht wieder. Einmal fuer den ganzen Bereich,
    ' nicht je Zelle - sonst kostet es 6.986 zusaetzliche Schreibvorgaenge.
    If letzteDaten >= INV_DATEN_START Then
        wsI.Range(wsI.Cells(INV_DATEN_START, 2), wsI.Cells(letzteDaten, 2)).NumberFormat = "0"
    End If

    ' Noch nicht gezaehlte Zeilen einfaerben (leeres GEZAEHLT-Feld).
    ' 16.08.: Die Formel-Variante (xlExpression mit AND) hat Excel hier abgelehnt.
    ' Jetzt die eingebaute Regel "leere Zellen" auf Spalte G - die braucht keine
    ' Formel und ist damit unabhaengig von der Sprachfassung.
    ' Bleibt nur Optik: scheitert es wieder, laeuft das Befuellen trotzdem durch.
    Dim fcOk As Boolean: fcOk = True
    If letzteDaten >= INV_DATEN_START Then
        On Error Resume Next
        With wsI.Range(wsI.Cells(INV_DATEN_START, 7), wsI.Cells(letzteDaten, 7))
            .FormatConditions.Delete
            With .FormatConditions.Add(Type:=xlBlanksCondition)
                .Interior.Color = RGB(255, 242, 204)
                .Font.Color = RGB(140, 100, 0)
            End With
        End With
        If Err.Number <> 0 Then fcOk = False: Err.Clear
        On Error GoTo 0
    End If

    ' Summenzeile
    Dim sumRow As Long: sumRow = sRow + 1
    wsI.Range("A" & sumRow & ":L" & sumRow).Interior.Color = RGB(31, 56, 100)
    wsI.Range("A" & sumRow & ":L" & sumRow).Font.Color = RGB(255, 255, 255)
    wsI.Range("A" & sumRow & ":L" & sumRow).Font.Bold = True
    wsI.Range("A" & sumRow & ":H" & sumRow).Merge
    wsI.Cells(sumRow, 1).Value = "SUMME EK-WERT (gezaehlt / gesamt):"
    wsI.Cells(sumRow, 1).HorizontalAlignment = xlRight
    wsI.Cells(sumRow, 9).Formula = "=SUM(I" & INV_DATEN_START & ":I" & letzteDaten & ")"
    wsI.Cells(sumRow, 9).NumberFormat = "#,##0.00 " & ChrW(8364)
    wsI.Rows(sumRow).RowHeight = 22

    Inventur_FortschrittAnzeigen wsI

BefuellenAufraeumen:
    Dim fehlerB As String: fehlerB = ""
    If Err.Number <> 0 Then fehlerB = Chr(10) & Chr(10) & "ABBRUCH mit Fehler: " & Err.Description
    On Error Resume Next
    Application.EnableEvents = True
    Application.Calculation = altCalcB
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Inventur geladen: " & (nr - 1) & " Artikel." & Chr(10) & Chr(10) & _
           IIf(fcOk, "Alle Zeilen stehen auf 'noch nicht gezaehlt' (gelb hinterlegt).", _
                     "Hinweis: Die gelbe Markierung liess sich nicht setzen." & Chr(10) & _
                     "Was offen ist, zeigt der Knopf NICHT GEZAEHLT.") & fehlerB, _
           IIf(fehlerB = "", vbInformation, vbCritical)
End Sub

' ================================================================
'  INVENTUR - BESTAENDE UEBERNEHMEN
' ================================================================
Sub Inventur_BestaendeUebernehmen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    Dim wsZ As Worksheet: Set wsZ = GetSheet("Abg")
    Dim wsB As Worksheet: Set wsB = GetSheet("Best")
    If wsA Is Nothing Or wsI Is Nothing Then Exit Sub
    If wsZ Is Nothing Then
        MsgBox "Sheet 'Zu- und Abgaenge' nicht gefunden - ohne Buchungsnachweis " & _
               "wird nichts uebernommen.", vbCritical
        Exit Sub
    End If

    Dim colAnz As Long: colAnz = Spalte_Finden(wsA, "ANZAHL")
    Dim colArt As Long: colArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim colEAN As Long: colEAN = Spalte_Finden(wsA, "EAN13")
    Dim colNr  As Long: colNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim colLag As Long: colLag = Spalte_Finden(wsA, "LAGERORT")
    Dim colVK  As Long: colVK = Spalte_Finden(wsA, "VK-PREIS")
    If colAnz = 0 Or colArt = 0 Then
        MsgBox "Spalten ANZAHL/ARTIKEL im Artikel-Sheet nicht gefunden.", vbCritical
        Exit Sub
    End If

    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If lastI < INV_DATEN_START Then MsgBox "Die Inventurliste ist leer.", vbInformation: Exit Sub

    ' ---- Durchlauf 1: nur zaehlen, nichts anfassen ----
    Dim nGezaehlt As Long, nAendert As Long, nOhneVerweis As Long, nVerschoben As Long
    Dim i As Long, az As Long
    For i = INV_DATEN_START To lastI
        If Trim(CStr(wsI.Cells(i, 7).Value)) <> "" And Trim(CStr(wsI.Cells(i, 3).Value)) <> "" Then
            nGezaehlt = nGezaehlt + 1
            az = val(wsI.Cells(i, INV_COL_ZEILE).Value)
            If az < ART_DATEN_START Then
                nOhneVerweis = nOhneVerweis + 1
            ElseIf LCase(Trim(CStr(wsA.Cells(az, colArt).Value))) <> LCase(Trim(CStr(wsI.Cells(i, 3).Value))) Then
                ' Der Verweis zeigt nicht mehr auf denselben Artikel - im Artikel-Sheet
                ' wurden Zeilen eingefuegt oder geloescht. Nicht schreiben.
                nVerschoben = nVerschoben + 1
            ElseIf val(wsA.Cells(az, colAnz).Value) <> val(wsI.Cells(i, 7).Value) Then
                nAendert = nAendert + 1
            End If
        End If
    Next i

    If nGezaehlt = 0 Then
        MsgBox "In der Liste ist noch nichts gezaehlt.", vbInformation
        Exit Sub
    End If

    Dim warn As String: warn = ""
    If nOhneVerweis > 0 Then warn = warn & Chr(10) & "! " & nOhneVerweis & _
        " Zeilen ohne Artikel-Verweis - werden UEBERSPRUNGEN (Liste neu befuellen)."
    If nVerschoben > 0 Then warn = warn & Chr(10) & "! " & nVerschoben & _
        " Zeilen zeigen auf einen anderen Artikel - werden UEBERSPRUNGEN."

    If MsgBox("Gezaehlt: " & nGezaehlt & " Artikel." & Chr(10) & _
              "Davon aendern sich " & nAendert & " Bestaende." & Chr(10) & _
              (lastI - INV_DATEN_START + 1 - nGezaehlt) & " nicht gezaehlte Zeilen bleiben unberuehrt." & _
              Chr(10) & warn & Chr(10) & Chr(10) & _
              "Jede Aenderung wird in 'Zu- und Abgaenge' gebucht und der alte Bestand " & _
              "in Spalte N gesichert." & Chr(10) & Chr(10) & _
              "Jetzt uebernehmen?", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Bestaende uebernehmen") = vbNo Then Exit Sub

    ' ---- Bestaende-Sheet einmal einlesen (ArtikelNr -> Zeile) ----
    ' Ohne diese Tabelle waere je Artikel eine Suche ueber 6.985 Zeilen noetig.
    Dim bIdx As Object: Set bIdx = CreateObject("Scripting.Dictionary")
    If Not wsB Is Nothing Then
        Dim lastB As Long: lastB = wsB.Cells(wsB.Rows.count, 2).End(xlUp).Row
        Dim bv As Variant
        If lastB >= 2 Then
            ' Eine Zeile mehr lesen: bei nur einer Zelle liefert .Value keinen Array,
            ' und UBound() wuerde abbrechen. Die Leerzeile faellt unten sowieso raus.
            bv = wsB.Range(wsB.Cells(2, 2), wsB.Cells(lastB + 1, 2)).Value
            Dim k As Long
            For k = 1 To UBound(bv, 1)
                Dim bs As String: bs = Trim(CStr(bv(k, 1)))
                If bs <> "" Then If Not bIdx.Exists(bs) Then bIdx.Add bs, k + 1
            Next k
        End If
    End If

    Dim altCalc As Long: altCalc = Application.Calculation
    On Error GoTo Aufraeumen
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' ---- Durchlauf 2: schreiben ----
    ' Jede Buchung wird SOFORT geschrieben, nicht gesammelt: bricht der Lauf mittendrin
    ' ab, sind sonst Bestaende geaendert, aber der Nachweis dazu fehlt.
    Dim updated As Long: updated = 0
    Dim jetzt As Date: jetzt = Now()
    Dim zRow As Long: zRow = wsZ.Cells(wsZ.Rows.count, 1).End(xlUp).Row + 1

    For i = INV_DATEN_START To lastI
        If Trim(CStr(wsI.Cells(i, 7).Value)) <> "" And Trim(CStr(wsI.Cells(i, 3).Value)) <> "" Then
            az = val(wsI.Cells(i, INV_COL_ZEILE).Value)
            If az >= ART_DATEN_START Then
                If LCase(Trim(CStr(wsA.Cells(az, colArt).Value))) = LCase(Trim(CStr(wsI.Cells(i, 3).Value))) Then
                    Dim vorher As Double: vorher = val(wsA.Cells(az, colAnz).Value)
                    Dim gezaehlt As Double: gezaehlt = val(wsI.Cells(i, 7).Value)
                    ' Alten Stand IMMER sichern, auch wenn sich nichts aendert
                    wsI.Cells(i, INV_COL_VORHER).Value = vorher
                    If vorher <> gezaehlt Then
                        wsA.Cells(az, colAnz).Value = gezaehlt
                        updated = updated + 1
                        ' Buchung sofort anhaengen
                        wsZ.Cells(zRow, 1).Value = jetzt
                        wsZ.Cells(zRow, 1).NumberFormat = "DD.MM.YYYY HH:MM"
                        If colEAN > 0 Then wsZ.Cells(zRow, 2).Value = wsA.Cells(az, colEAN).Value
                        If colNr > 0 Then wsZ.Cells(zRow, 3).Value = wsA.Cells(az, colNr).Value
                        wsZ.Cells(zRow, 4).Value = wsA.Cells(az, colArt).Value
                        wsZ.Cells(zRow, 5).Value = Abs(gezaehlt - vorher)
                        wsZ.Cells(zRow, 6).Value = IIf(gezaehlt > vorher, "Inventur Zugang", "Inventur Abgang")
                        If colLag > 0 Then wsZ.Cells(zRow, 7).Value = wsA.Cells(az, colLag).Value
                        wsZ.Cells(zRow, 8).Value = BENUTZER
                        wsZ.Cells(zRow, 9).Value = "Inventur: vorher " & Format(vorher, "0") & _
                                                   ", gezaehlt " & Format(gezaehlt, "0")
                        zRow = zRow + 1
                        ' Bestaende-Sheet mitziehen (sonst laufen beide auseinander)
                        If Not wsB Is Nothing And colNr > 0 Then
                            Dim schl As String: schl = Trim(CStr(wsA.Cells(az, colNr).Value))
                            If schl <> "" Then
                                If bIdx.Exists(schl) Then
                                    Dim bz As Long: bz = bIdx(schl)
                                    wsB.Cells(bz, 4).Value = gezaehlt
                                    If colVK > 0 Then wsB.Cells(bz, 6).Value = _
                                        Round(gezaehlt * val(wsA.Cells(az, colVK).Value), 2)
                                    wsB.Cells(bz, 10).Value = IIf(gezaehlt = 0, "! Nachbestellung", "OK")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

Aufraeumen:
    ' Fehlerlage ZUERST sichern, dann erst aufraeumen: ein Fehler im Aufraeumen selbst
    ' wuerde sonst wieder hierher springen.
    Dim Fehler As String: Fehler = ""
    If Err.Number <> 0 Then Fehler = Chr(10) & Chr(10) & "ABBRUCH mit Fehler: " & Err.Description
    On Error Resume Next
    Application.EnableEvents = True
    Application.Calculation = altCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0

    MsgBox updated & " Bestaende geaendert und gebucht." & Chr(10) & _
           (nGezaehlt - updated) & " gezaehlte Artikel waren bereits richtig." & Chr(10) & _
           "Alter Stand steht in Spalte N (versteckt)." & warn & Chr(10) & Chr(10) & _
           "Die Schnellansicht wird NICHT automatisch neu gebaut - " & _
           "dafuer den Knopf SCHNELLANSICHT im Artikel-Blatt druecken." & Fehler, _
           IIf(Fehler = "", vbInformation, vbCritical)
End Sub

' ================================================================
'  INVENTUR - WAS FEHLT NOCH? (Knopf K2/L2)
' ================================================================
Sub Inventur_NichtGezaehlt()
    Dim wsI As Worksheet: Set wsI = GetSheet("Inventur")
    If wsI Is Nothing Then Exit Sub

    Dim lastI As Long: lastI = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
    If lastI < INV_DATEN_START Then MsgBox "Die Inventurliste ist leer.", vbInformation: Exit Sub

    ' Filter an? Dann ausschalten und alles wieder zeigen.
    If wsI.FilterMode Or wsI.AutoFilterMode Then
        wsI.AutoFilterMode = False
        MsgBox "Filter aus - es werden wieder alle Zeilen gezeigt.", vbInformation
        Exit Sub
    End If

    ' Spalten C bis G in einem Zug lesen - Zelle fuer Zelle waere bei 6.986 Zeilen traege
    Dim dat As Variant
    dat = wsI.Range(wsI.Cells(INV_DATEN_START, 3), wsI.Cells(lastI, 7)).Value
    Dim offen As Long: offen = 0
    Dim i As Long
    For i = 1 To UBound(dat, 1)
        If Trim(CStr(dat(i, 1))) <> "" And Trim(CStr(dat(i, 5))) = "" Then offen = offen + 1
    Next i

    Inventur_FortschrittAnzeigen wsI

    Dim gesamt As Long: gesamt = lastI - INV_DATEN_START + 1
    If offen = 0 Then
        MsgBox "Alle " & gesamt & " Artikel sind gezaehlt.", vbInformation
        Exit Sub
    End If

    ' Nur die offenen Zeilen zeigen (Summenzeile bleibt ausserhalb des Filterbereichs)
    wsI.Range(wsI.Cells(5, 1), wsI.Cells(lastI, INV_COL_WER)).AutoFilter Field:=7, Criteria1:="="

    MsgBox "Noch nicht gezaehlt: " & offen & " von " & gesamt & " Artikeln." & Chr(10) & _
           "Gezaehlt: " & (gesamt - offen) & Chr(10) & Chr(10) & _
           "Es werden jetzt nur die offenen Zeilen gezeigt." & Chr(10) & _
           "Nochmal auf NICHT GEZAEHLT klicken zeigt wieder alle.", vbInformation
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

    Dim lastRow As Long: lastRow = wsIS.Cells(wsIS.Rows.count, 3).End(xlUp).Row
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
    Dim nurZahlen As Boolean: nurZahlen = (such = CStr(val(such)) And val(such) > 0)
    Dim lastA As Long: lastA = LetzteZeile(wsA, colArt)
    Dim sRow As Long: sRow = 4
    Dim nr As Long: nr = 1
    Dim hellgrauS As Long: hellgrauS = RGB(242, 242, 242)
    Dim i As Long, w As Integer, passt As Boolean, suchIn As String

    For i = ART_DATEN_START To lastA
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
    Dim lastRow As Long: lastRow = wsIS.Cells(wsIS.Rows.count, 3).End(xlUp).Row
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

    g_InvSucheArtikelZeile = val(wsIS.Cells(listZeile, 8).Value)
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
    Dim lastRow As Long: lastRow = wsIS.Cells(wsIS.Rows.count, 3).End(xlUp).Row
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

    Dim menge As Double: menge = val(mengeStr)
    Dim az As Long: az = g_InvSucheArtikelZeile

    If Not wsI Is Nothing Then
        Dim colEAN_I As Long: colEAN_I = Spalte_Finden(wsA, "EAN13")
        Dim colArt_I As Long: colArt_I = Spalte_Finden(wsA, "ARTIKEL")
        Dim colLag_I As Long: colLag_I = Spalte_Finden(wsA, "LAGERORT")
        Dim colEK_I  As Long: colEK_I = Spalte_Finden(wsA, "EK-PREIS")
        Dim colAnz_I As Long: colAnz_I = Spalte_Finden(wsA, "ANZAHL")
        Dim lastInv As Long: lastInv = wsI.Cells(wsI.Rows.count, 3).End(xlUp).Row
        ' Zuordnung ueber den Zeilenverweis (Spalte M), nicht ueber den Namen -
        ' 273 Artikelnamen sind mehrfach vergeben.
        Dim gefunden As Boolean: gefunden = False
        Dim ii As Long
        For ii = INV_DATEN_START To lastInv
            If val(wsI.Cells(ii, INV_COL_ZEILE).Value) = az Then
                wsI.Cells(ii, 7).Value = menge
                Inventur_BelegSchreiben wsI, ii
                gefunden = True
                Exit For
            End If
        Next ii
        If Not gefunden Then
            Dim nRow As Long
            If lastInv < INV_DATEN_START Then nRow = INV_DATEN_START Else nRow = lastInv + 1
            ' Nicht in die Summenzeile schreiben
            Dim sumRowS As Long: sumRowS = Inventur_SummenZeile(wsI)
            If sumRowS > 0 And nRow >= sumRowS Then
                wsI.Rows(sumRowS).Insert Shift:=xlDown
                nRow = sumRowS
                sumRowS = sumRowS + 1
            End If
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
            wsI.Cells(nRow, INV_COL_ZEILE).Value = az
            Inventur_BelegSchreiben wsI, nRow
            If sumRowS > 0 Then
                wsI.Cells(sumRowS, 9).Formula = "=SUM(I" & INV_DATEN_START & ":I" & (sumRowS - 1) & ")"
            End If
        End If
    End If

    If Not wsI Is Nothing Then Inventur_FortschrittAnzeigen wsI
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
    MsgBox s, vbInformation, "Spaltendiagnose"
End Sub

' ================================================================
'  ARTIKEL-ZEILE MARKIEREN (aufgerufen aus Worksheet_SelectionChange)
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
'  ARTIKEL SUCHEN (Button-Wrapper, liest Suchfeld B2)
' ================================================================
Sub Artikel_Suchen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Dim such As String: such = Trim(CStr(wsA.Cells(2, 2).Value))
    If such = "" Then
        Artikel_Suche_Leeren
        Exit Sub
    End If
    Application.ScreenUpdating = False

    ' --- Spalten dynamisch finden ---
    Dim cArt  As Long: cArt = Spalte_Finden(wsA, "ARTIKEL")
    Dim cEAN  As Long: cEAN = Spalte_Finden(wsA, "EAN")
    Dim cNr   As Long: cNr = Spalte_Finden(wsA, "ARTIKELNR")
    Dim cLag  As Long: cLag = Spalte_Finden(wsA, "LAGERORT")
    Dim cWgr  As Long: cWgr = Spalte_Finden(wsA, "WARENGRUPPE")
    If cArt = 0 Then cArt = Spalte_Finden(wsA, "Artikel")
    If cArt = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Spalte 'Artikel' nicht gefunden!", vbExclamation
        Exit Sub
    End If

    ' --- Erst ALLE Zeilen einblenden, dann lastRow berechnen ---
    wsA.Rows("5:50000").Hidden = False
    If wsA.AutoFilterMode Then wsA.AutoFilterMode = False
    Dim lastRow As Long: lastRow = LetzteZeile(wsA, cArt)
    If lastRow < 5 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' --- Suchmodus: Zahlen = EAN/ArtikelNr, Text = nur Artikelname ---
    Dim nurZahlen As Boolean: nurZahlen = IsNumeric(such) And Len(such) > 0
    Dim woerter() As String: woerter = Split(LCase(such), " ")
    Dim treffer As Long: treffer = 0
    Dim i As Long, w As Integer, passt As Boolean, suchIn As String

    For i = 5 To lastRow
        If Trim(CStr(wsA.Cells(i, cArt).Value)) = "" Then
            wsA.Rows(i).Hidden = True
        Else
            If nurZahlen Then
                ' Zahlensuche: EAN + Artikelnummer
                suchIn = ""
                If cEAN > 0 Then suchIn = suchIn & " " & LCase(CStr(wsA.Cells(i, cEAN).Value))
                If cNr > 0 Then suchIn = suchIn & " " & LCase(CStr(wsA.Cells(i, cNr).Value))
            Else
                ' Textsuche: NUR Artikelname (keine Lagerort/Warengruppe-Falstreffer)
                suchIn = LCase(CStr(wsA.Cells(i, cArt).Value))
            End If

            passt = True
            For w = 0 To UBound(woerter)
                If Trim(woerter(w)) <> "" Then
                    If InStr(suchIn, Trim(woerter(w))) = 0 Then
                        passt = False: Exit For
                    End If
                End If
            Next w
            wsA.Rows(i).Hidden = Not passt
            If passt Then treffer = treffer + 1
        End If
    Next i

    ' Trefferzahl im Shape anzeigen
    On Error Resume Next
    wsA.Shapes("trefferAnzeige").TextFrame.Characters.text = treffer & " Treffer"
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  ARTIKEL GESAMTANZAHL (zeigt alle Artikel im trefferAnzeige-Shape)
' ================================================================
Sub Artikel_Anzahl_Anzeigen()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Dim cArt As Long: cArt = Spalte_Finden(wsA, "ARTIKEL")
    If cArt = 0 Then Exit Sub
    wsA.Rows("5:50000").Hidden = False
    If wsA.AutoFilterMode Then wsA.AutoFilterMode = False
    Dim lastRow As Long: lastRow = LetzteZeile(wsA, cArt)
    Dim gesamt As Long: gesamt = 0
    Dim i As Long
    For i = 5 To lastRow
        If Trim(CStr(wsA.Cells(i, cArt).Value)) <> "" Then gesamt = gesamt + 1
    Next i
    On Error Resume Next
    wsA.Shapes("trefferAnzeige").TextFrame.Characters.text = gesamt & " Artikel"
    On Error GoTo 0
End Sub

' ================================================================
'  ARTIKEL SUCHE LEEREN
' ================================================================
Sub Artikel_Suche_Leeren()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Application.EnableEvents = False
    wsA.Cells(2, 2).Value = ""
    Application.EnableEvents = True
    Artikel_Anzahl_Anzeigen
End Sub

' ================================================================
'  ARTIKEL AKTUALISIEREN (Sheet neu laden / Anzeige auffrischen)
' ================================================================
Sub Artikel_Aktualisieren()
    Dim wsA As Worksheet: Set wsA = GetSheet("Artikel")
    If wsA Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wsA.Cells(2, 2).Value = ""
    Application.EnableEvents = True
    Artikel_Anzahl_Anzeigen
    Application.ScreenUpdating = True
    Application.StatusBar = "Artikel aktualisiert."
End Sub
