# Lagerverwaltung Excel – Claude-Referenz

> Diese Datei wird von Claude automatisch beim Session-Start eingelesen.
> Details: `PROJEKTINFO.md` · Offene Fehler: `Pruefbericht_Lagerverwaltung_2026-06-12.md`

---

## Was ist das?
Excel-basierte Lagerverwaltung (FB300, Braun). Reines VBA, keine Frameworks.
Alle Makros leben in `2026_Lagerverwaltung_V3.xlsm` – die `.bas`-Dateien im Ordner sind
nur **Exporte** (erzeugt vom GITHUB-Button). Maßgeblich ist immer der Stand in der Mappe.

## Dateien & Pfade
| Datei | Zweck |
|---|---|
| `2026_Lagerverwaltung_V3.xlsm` | Haupt-Datei (Master, enthält alle Module) |
| `LagerMakros.bas` | Kern: Suche, ZuAbgang, Etikett (Zebra GK420d), Inventur, Setup_Ausfuehren |
| `NeueModule.bas` | Toolbar, GitHub_Export (AKTIV), ArtikelDetail, SchnellDetail, BewPopup, Inv-Sheets |
| `NeuArtikelModul.bas` | Neuer-Artikel-Formular, Suche & Scanner, EAN-Generator (Prefix 200) |
| `ArtikelFix.bas` | ⚠ DEFEKT/abgeschnitten (endet Z. 286) – NICHT importieren, frisch exportieren! |
| `DuplikatPruefer.bas` | Duplikat-Prüfung – ⚠ passt nicht zum V3-Layout (Header in Z. 2 erwartet) |
| `CSV_Import.bas` | Import MF_DACH_MAT.csv (Semikolon, 37 Felder, VK = EK × 1,35) |
| `GitHub_Export_NEU.bas` | ⚠ VERALTET (hartkodierter alter Pfad) – nicht verwenden |
| `lager.json` | Export der Schnellansicht – ⚠ aktuell Datenmüll (falsches Spalten-Mapping) |
| `PROJEKTINFO.md` | Ausführliche Doku – ⚠ Pfade teils veraltet (alter Repo-Ordner) |
| `Backup\` | Backups (im Windows-Explorer anlegen!) |
| `Prüfen bevor löschen\Backup\Lagerverw. FB300\` | ⚠ Enthält das EINZIGE Git-Repo – nicht löschen bis Umzug! |

## Git / GitHub
- Remote: `https://github.com/ddmbraun/Lagerverwaltung-FB300.git` (Branch main) – **ÖFFENTLICH!**
- Repo enthält GitHub-Pages-Webansicht (`index.html` liest `lager.json`). EK-Preise im Export: von Frank am 12.06. ausdrücklich gewünscht (öffentlich sichtbar).
- **Neues Konzept (12.06.):** `.git` liegt direkt in `07_Lagerverwaltung-Excel\` (= Repo-Root = Mappen-Ordner).
  `GitHub_Export` nutzt GIT_DIR = `ThisWorkbook.Path` + Sicherheitscheck auf `.git`-Ordner.
  `.gitignore` schließt Backups, „Prüfen bevor löschen" und MF_DACH_MAT.csv aus.
- Letzter erfolgreicher Push: 07.06.2026. Im alten Backup-Repo liegen uncommittete Änderungen (werden durch neuen Stand ersetzt).

## Artikel-Sheet V3-Layout (WICHTIG)
| Zeile | Inhalt |
|---|---|
| 1 | Titel „Artikel - Lagerverwaltung" |
| 2 | Suchfeld B2 + SUCHEN/LEEREN/AKTUALISIEREN + trefferAnzeige |
| 3 | Toolbar: GITHUB, NEUER ARTIKEL, ZU-/ABGANG, ETIKETT, EK ausbl., FILTER LOESCHEN, SCHNELLANSICHT |
| 4 | Spaltenüberschriften (EAN13, ARTIKEL, VK-PREIS, …) |
| 5+ | Artikeldaten |

`Spalte_Finden(ws, header)` sucht beim Artikel-Sheet ab Zeile 3 (Titel-Falle).
⚠ Viele Leseschleifen starten noch bei `i = 3` → Headerzeile 4 wird als Artikel mitgelesen (offener Fehler H2).

## Schnellansicht – Layout-Konflikt (offener Fehler H1/K2)
`Schnellansicht_Aktualisieren` schreibt A=EAN, B=Artikel, C=VK, D=Anzahl, E=EK, F=Zeilenverweis (ab Z. 4).
`ExportLagerJSON`, `EK_Toggle`, Such-Popup und Doppelklick-Handler erwarten andere Layouts.
→ Vor jeder Änderung an Schnellansicht-Funktionen den Prüfbericht lesen.

## Standard-Vorgehen bei Modul-Änderungen
1. Backup der `.xlsm` im **Windows-Explorer** nach `Backup\` (NIE per Sandbox-Shell-cp!)
2. `.bas` mit Edit-Tool ändern
3. In Excel: Alt+F11 → altes Modul entfernen (Nein = nicht exportieren) → Datei → Importieren
4. Debuggen → Kompilieren (Syntaxcheck)
5. Bei Struktur-/Event-Änderungen: `Setup_Ausfuehren` einmal ausführen
   (braucht Trust Center: „Zugriff auf VBA-Projektobjektmodell vertrauen" = AN)
6. Importreihenfolge Vollimport: LagerMakros → NeueModule → NeuArtikelModul → Setup_Ausfuehren

## Bekannte Fallen (gleiche wie PAM/HV)
- Sandbox-Shell kann veraltete/abgeschnittene Kopien des D:-Mounts zeigen.
  **Maßgeblich ist das Datei-Werkzeug (Read/Edit).** Keine Edits/Backups/Kopien per Shell.
- Backups NICHT per Shell-`cp` → Windows-Explorer.
- `git add -A` in GitHub_Export niemals laufen lassen, solange GIT_DIR nicht verifiziert ist
  (Gefahr: committet bei falschem Pfad fremde Projektordner).
- ArtikelFix.bas im Ordner ist abgeschnitten – Stand in der Mappe ist der gute.

## Offene Befunde (Stand 12.06.2026)
→ Vollständig mit Lösungen: `Pruefbericht_Lagerverwaltung_2026-06-12.md`

| Nr | Befund | Status |
|---|---|---|
| K1 | GitHub-Upload defekt (Repo im Lösch-Ordner, GIT_DIR falsch) | 🟢 VBA gefixt – Frank muss noch: `.git` im Explorer verschieben + Modul importieren |
| K2 | lager.json Spalten-Mapping falsch, EK landet im Feld „ean", kein UTF-8 | 🟢 ExportLagerJSON neu (Artikel-Sheet, ab Z. 5, UTF-8) – Import + Testlauf offen |
| K3 | ArtikelFix.bas abgeschnitten | ⬜ offen – wird beim nächsten GITHUB-Klick automatisch frisch exportiert |
| H1 | Schnellansicht: 3 verschiedene Spalten-Layouts | ⬜ offen |
| H2 | Leseschleifen ab Zeile 3 statt 5 (Header wird Artikel) | ⬜ offen |
| H3 | DuplikatPruefer inkompatibel mit V3-Layout | ⬜ offen |
| M1 | GitHub_Export_NEU.bas Altlast entfernen | ⬜ offen |
| M2 | CSV-Import überschreibt gepflegte VK-Preise | ⬜ offen |
| M3 | Keine Duplikatprüfung bei NeuerArtikel_Speichern | ⬜ offen |
| M4 | InvSuche-UserForm „Fehler beim Zugriff auf Pfad/Datei" | ⬜ offen |
| M5 | LagerMakros ohne Option Explicit | ⬜ offen |

## Wichtige Konstanten
```vba
' LagerMakros.bas
Const ZEBRA_DRUCKER = "ZDesigner GK420d"
Public Const BENUTZER = "Frank"
Const INV_DATEN_START = 6      ' Inventur: Daten ab Zeile 6
' EAN-Generator: Prefix 200 = intern, EAN13-Prüfziffer alternierend ×1/×3
' CSV_Import: VK_AUFSCHLAG = 1.35
```
