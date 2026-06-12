# Lagerverwaltung V3 – Projektdokumentation

> Dieses Dokument am Anfang eines neuen Chats einfügen, damit Claude den vollen Kontext hat.

---

## Übersicht

Excel-basierte Lagerverwaltung für FB300 (Braun).
Single-File VBA-Anwendung, keine Frameworks.
Basis: `2026_Lagerverwaltung_KOMPLETT.xlsm` → umgebaut zu V3 mit neuem Toolbar-Design.

---

## Dateien & Pfade

| Datei | Zweck |
|---|---|
| `D:\_KI-Projekte-2026\07_Lagerverwaltung-Excel\Lagerverw. FB300\Lagerverwaltung_2026_V3\2026_Lagerverwaltung_V3.xlsm` | Haupt-Datei (Master) |
| `...\LagerMakros.bas` | Kern-Makros: Suche, ZuAbgang, Etikett, EK-Toggle, Inventur, Schnellansicht |
| `...\NeueModule.bas` | Neue Sheets & Toolbar-Setup, GitHub-Export, ArtikelDetail, SchnellDetail, BewPopup |
| `...\NeuArtikelModul.bas` | Neuer-Artikel-Formular, Suche & Scanner Sheet |
| `...\ArtikelFix.bas` | Hilfsfunktionen Artikel-Korrekturen |
| `...\DuplikatPruefer.bas` | Duplikat-Prüfung Artikel |
| `...\GitHub_Export_NEU.bas` | Veraltet – GitHub-Export ist jetzt in NeueModule.bas |

**Git-Repository:**
- Lokaler Pfad: `D:\_KI-Projekte-2026\07_Lagerverwaltung-Excel\Lagerverw. FB300\`
- Remote: `https://github.com/ddmbraun/Lagerverwaltung-FB300.git`
- Branch: `main`

---

## Tabellenblätter

| Sheet | Zweck |
|---|---|
| `Schnellansicht` | Hauptansicht mit Artikel-Liste und Suche |
| `Artikel` | Vollständige Artikelverwaltung mit Toolbar |
| `Bestände` | Lagerbestände |
| `Zu- und Abgänge` | Buchungshistorie |
| `Warengruppen` | Dropdown-Daten |
| `Lagerorte` | Dropdown-Daten |
| `Dashboard` | Übersicht |
| `Inventurliste` | Inventur-Erfassung |
| `InvSuche` | Inventur-Suchmaske (per Setup erstellt) |
| `InvDaten` | Inventur-Daten (per Setup erstellt) |
| `InvEingabe` | Inventur-Eingabe (per Setup erstellt) |
| `ArtikelDetail` | Artikel-Detailansicht (per Setup erstellt, Doppelklick in Artikel) |
| `BewPopup` | Bewegungs-Popup (Bestandsverlauf) |
| `SchnellDetail` | Detail-Popup aus Schnellansicht |

---

## Artikel-Sheet Struktur (WICHTIG)

Nach `Setup_Ausfuehren` hat das Artikel-Sheet dieses Layout:

| Zeile | Inhalt |
|---|---|
| 1 | Titelleiste „Artikel - Lagerverwaltung" (dunkelblau, merged A1:U1) |
| 2 | Suchfeld (B2, gelb) + Buttons: SUCHEN, LEEREN, AKTUALISIEREN, trefferAnzeige |
| 3 | Toolbar-Buttons: GITHUB, NEUER ARTIKEL, ZU-/ABGANG, ETIKETT, EK ausbl., FILTER LOESCHEN, SCHNELLANSICHT |
| 4 | Spaltenüberschriften (EAN13, ARTIKEL, VK-PREIS, EK-PREIS, ANZAHL, ...) |
| 5+ | Artikeldaten |

**Kritisch:** `Spalte_Finden()` sucht für das Artikel-Sheet ab Zeile 3 (nicht ab Zeile 1), weil Zeile 1 den Titel „Artikel - Lagerverwaltung" enthält, der sonst fälschlich als „ARTIKEL"-Spalte erkannt wird.

---

## Module & wichtige Funktionen

### LagerMakros.bas

| Funktion | Beschreibung |
|---|---|
| `Setup_Ausfuehren()` | Initialisiert alles: Toolbar, Sheets, Events. Einmal ausführen nach Modul-Import |
| `Spalte_Finden(ws, headerName)` | Sucht Spalte per Überschrift. Für Artikel-Sheet: Start ab Zeile 3 |
| `GetSheet(name)` | Sheet per Teilname finden (robust gegen Umlaute) |
| `Artikel_Suchen()` | Suche in Artikel-Sheet. Text → nur Artikelname. Zahlen → EAN + Artikelnummer |
| `Artikel_Suche_Leeren()` | Suchfeld leeren + alle Zeilen einblenden + Gesamtanzahl anzeigen |
| `Artikel_Aktualisieren()` | Sheet-Reset + Gesamtanzahl anzeigen |
| `Artikel_Anzahl_Anzeigen()` | Zählt alle Artikel und zeigt im trefferAnzeige-Shape an |
| `Artikel_Zeile_Markieren(Target)` | Angeklickte Artikelzeile gelb markieren (ab Zeile 5, ab Spalte 2) |
| `Filter_Loeschen()` | Alle Zeilen einblenden + Suchfeld leeren + Anzahl anzeigen |
| `EK_Toggle()` | EK-Preis Spalte ein-/ausblenden |
| `ZuAbgang_Buchen()` | Zu-/Abgangsbuchung für markierten Artikel |
| `Etikett_Drucken()` | Zebra-Etikett drucken (Drucker: ZDesigner GK420d) |
| `NeuerArtikel()` | Neuen Artikel anlegen |

### NeueModule.bas

| Funktion | Beschreibung |
|---|---|
| `Setup_Artikel_Toolbar()` | Erstellt Toolbar (Zeilen 1-3) + installiert Events ins Artikel-Sheet |
| `GitHub_Export()` | Exportiert .bas-Dateien + lager.json → Git commit + push |
| `Schnellansicht_Oeffnen()` | Wechselt zum Schnellansicht-Sheet |
| `ArtikelDetail_Laden(zeile)` | Lädt ArtikelDetail-Sheet (bei Doppelklick in Artikel) |
| `SchnellDetail_Laden(ean)` | Lädt SchnellDetail-Sheet |

### NeuArtikelModul.bas

| Funktion | Beschreibung |
|---|---|
| `EK_Toggle()` | Wrapper → ruft `LagerMakros.EK_Toggle` auf |
| `Artikel_FilterLoeschen()` | Filter leeren (alte Implementierung, nicht mehr Haupt-Button) |
| `Artikel_Suchen(wsA, such)` | Suche mit Parametern (aus V2, wird intern verwendet) |
| `AllesNeuEinrichten()` | Erstellt Suche & Scanner + Neuer Artikel Sheet komplett neu |
| `Setup_NeuerArtikel()` | Erstellt „Neuer Artikel"-Sheet (Felder, Buttons, Events, versteckt) |
| `DropdownsEinrichten()` | Dropdowns für B8/B9/B10/B13 einrichten. Attribut-Werte in Spalte G (Hidden) |
| `NeuerArtikel_Oeffnen()` | Sheet sichtbar machen + Felder leeren (Button-Aktion) |
| `NeuerArtikel_Speichern()` | Felder in Artikel-Sheet + Bestände schreiben |
| `NeuerArtikel_FelderLeeren()` | B4:B18 + D5:D25 leeren, Defaults setzen (B10=19, B11=Stk) |
| `NeuerArtikel_VK_Berechnen(ws)` | VK = EK × (1 + Aufschlag%) × (1 + MwSt%). Trigger: B6, B18, B10 |
| `NeuerArtikel_EAN_Generieren(ws)` | Interne EAN13 (Prefix 200) generieren. Trigger: Doppelklick auf **B12** |
| `NeuerArtikel_Events_Jetzt()` | Events in „Neuer Artikel" reinstallieren (Alt+F8 → ausführen wenn Buttons tot) |
| `NeuerArtikel_QuickFix()` | Alles auf einmal: Spalten E-G verstecken + Dropdowns + Events reinstallieren |
| `NeuerArtikel_EventCode()` | String-Funktion: liefert den Event-Code für das Sheet (Change + BeforeDoubleClick + Deactivate) |
| `NeuerArtikel_Vorschlaege(Target)` | Tippt in B8/B9/B10/B13 → zeigt Vorschläge in Spalte D |
| `NeuerArtikel_VorschlagUebernehmen(Target)` | Doppelklick auf Vorschlag in D → übernimmt in passendes Feld |

---

## Suche im Artikel-Sheet

**Textsuche** (z.B. „Bohrer 6"):
- Sucht NUR in der ARTIKEL-Spalte (Artikelname)
- Mehrwort-Suche: alle Wörter müssen im Namen vorkommen (AND-Logik)
- Beispiel: „Bohrer 6" → findet nur Artikel wo BEIDE Wörter im Namen stehen

**Zahlsuche** (z.B. EAN oder Artikelnummer):
- Erkennung: `IsNumeric(such)` = True
- Sucht in EAN-Spalte + Artikelnummer-Spalte
- Kein Mindest-Stellenlimit (früher: 8 Stellen)

**Gesamtanzahl:**
- Wird automatisch angezeigt wenn Suchfeld leer / beim Öffnen des Sheets
- Shape: `trefferAnzeige` (Zeile 2, neben SUCHEN-Button)

---

## GitHub-Button (GITHUB in Toolbar)

- Exportiert alle .bas-Module in den V3-Ordner
- Exportiert `lager.json` (Schnellansicht-Daten)
- Führt `git add -A → commit → push` aus
- **Dynamische Pfade:** Git-Root = Parent-Ordner von `ThisWorkbook.Path`
  - Funktioniert solange der V3-Ordner innerhalb des Git-Repos liegt
  - Git-Root: `D:\_KI-Projekte-2026\07_Lagerverwaltung-Excel\Lagerverw. FB300\`

---

## Module importieren (Standard-Vorgehen)

Nach jeder Änderung an .bas-Dateien:

1. Excel → `Alt+F11` → VBA-Editor
2. Für jedes geänderte Modul: Rechtsklick → **Entfernen** → **Nein** (nicht exportieren)
3. Menü **Datei → Importieren** → entsprechende .bas-Datei wählen
4. VBA-Editor schließen
5. Bei Strukturänderungen (Toolbar/Events): `Setup_Ausfuehren` einmal aufrufen

**Reihenfolge beim Vollimport (alle 3 Module):**
`LagerMakros.bas` → `NeueModule.bas` → `NeuArtikelModul.bas` → dann `Setup_Ausfuehren`

---

## Bekannte Probleme & Status

## Neuer Artikel Sheet – Architektur (Stand 2026-06-07)

| Zeile | Inhalt |
|---|---|
| 1 | Titel „NEUEN ARTIKEL ANLEGEN" (dunkelblau) |
| 2 | Hinweis-Text (rot, kursiv) |
| 3 | Leerzeile (8px) |
| 4–17 | Eingabefelder: Label (A), gelbes Feld (B), Hinweis (C) |
| 18 | Aufschlag % → berechnet VK automatisch |
| 19 | Leerzeile |
| 20 | SPEICHERN (A20, grün) · FELDER LEEREN (B20, rot) – je Doppelklick |

**Feldliste B4–B18:** Artikelnr, Artikel, EK, VK, Warengruppe, Lagerort, MwSt (19), Einheit (Stk), EAN13, Attribut, TextA, TextB, Anfangsbestand, Vermerk, Aufschlag%

**Events (Worksheet-Code im Sheet):**
- `Worksheet_Change`: B6/B18/B10 → VK_Berechnen; B8/B9/B10/B13 → Vorschläge in D
- `Worksheet_BeforeDoubleClick`: A20=Speichern, B20=FelderLeeren, **B12=EAN generieren**, D5:D25=Vorschlag übernehmen
- `Worksheet_Deactivate`: Sheet wird automatisch versteckt bei Tab-Wechsel

**VK-Formel:** `VK = Round(EK × (1 + Aufschlag/100) × (1 + MwSt/100), 2)`

**EAN-Generator:** Prefix 200 = intern. Sucht höchste vorhandene interne EAN im Artikel-Sheet, nächste Nummer +1, Prüfziffer nach EAN13-Standard (alternierend ×1/×3).
- Trigger: **Doppelklick auf B12** (gelbes EAN-Feld, wenn leer)

**Attribut-Dropdown:** Eindeutige Werte aus Artikel-Sheet Attribut-Spalte → gespeichert in Spalte G (Hidden=True), Validation auf B13 zeigt Dropdown.

**Spalten:** A=28, B=35, C=35, D=25, E–G=Hidden (Hilfsdaten)

---

## Bekannte Probleme & Status

| Problem | Status |
|---|---|
| Toolbar-Buttons nicht sichtbar (V2→V3 Migration) | ✅ Behoben: `Setup_Artikel_Toolbar` wird jetzt aufgerufen |
| `EK_Toggle` → Laufzeitfehler 424 | ✅ Behoben: `NeuArtikelModul.EK_Toggle` rief `EK_Ausblenden` auf (existiert nicht), korrigiert auf `EK_Toggle` |
| Textsuche findet falsche Treffer (Lagerort/Warengruppe) | ✅ Behoben: Text sucht nur noch in Artikelname |
| `Spalte_Finden` findet Titel statt Header | ✅ Behoben: `startRow = 3` für Artikel-Sheet |
| EAN-Suche sehr langsam | ✅ Behoben: `lastRow` jetzt aus korrekter ARTIKEL-Spalte |
| Strikethrough im Suchfeld | ✅ Behoben: Explizit `Font.Strikethrough = False` in Setup |
| Zeile 3 (Buttons) wird gelb markiert | ✅ Behoben: `SelectionChange` nur ab Zeile 5 UND Spalte >= 2 |
| FILTER LOESCHEN zeigt unerwartete Spalte | ✅ Behoben: `Filter_Loeschen` neu implementiert |
| GitHub push schlägt fehl (Merge-Konflikt) | ✅ Behoben: `git checkout --ours` + commit + push manuell ausgeführt |
| InvSuche UserForm „Fehler beim Zugriff auf Pfad/Datei" | ⚠️ Offen: Nur Inventur-Suche betroffen, Rest funktioniert |
| EAN-Generator Doppelklick D12 funktionierte nicht | ✅ Behoben: Trigger auf B12 (Doppelklick auf gelbes EAN-Feld) verlegt. D12 wurde von FelderLeeren (D5:D25) gecleart |
| Spalte G mit Attribut-Werten sichtbar | ✅ Behoben: `Hidden = True` statt `ColumnWidth = 1` in `DropdownsEinrichten` |
| VK-Formel ohne MwSt | ✅ Behoben: `VK = EK × (1+Aufschlag%) × (1+MwSt%)`. Auch B10-Änderung triggert Neuberechnung |

---

## Noch nicht umgesetzt / geplant

- Schnellansicht-Suche überprüfen/anpassen (gleiche Spalten-Logik wie Artikel)
- InvSuche UserForm-Fehler lösen
- ZU-/ABGANG und ETIKETT-Buttons in V3 vollständig implementieren (aktuell aus KOMPLETT übernommen)

---

## Wichtige Konstanten

```vba
' LagerMakros.bas
Const ZEBRA_DRUCKER = "ZDesigner GK420d"
Const BENUTZER      = "Frank"
Const INV_DATEN_START = 6   ' Inventur: Daten ab Zeile 6

' NeueModule.bas (GitHub_Export)
' GIT_DIR  = Parent von ThisWorkbook.Path (dynamisch)
' EXPORT_DIR = ThisWorkbook.Path (dynamisch)
' Module: LagerMakros, NeueModule, NeuArtikelModul, ArtikelFix, DuplikatPruefer
```

---

## Trust Center Einstellung (einmalig)

Für programmatischen VBProject-Zugriff (Setup_Ausfuehren):
**Excel → Datei → Optionen → Trust Center → Trust Center-Einstellungen → Makroeinstellungen → „Zugriff auf das VBA-Projektobjektmodell vertrauen" → aktivieren**
