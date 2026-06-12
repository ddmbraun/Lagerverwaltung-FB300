# Prüfbericht Lagerverwaltung V3 – 12.06.2026

Nur-Lese-Prüfung aller VBA-Module, lager.json, CSV und Git-Status. **Keine Änderungen vorgenommen.**

Geprüft: LagerMakros.bas, NeueModule.bas, NeuArtikelModul.bas, ArtikelFix.bas, DuplikatPruefer.bas, CSV_Import.bas, GitHub_Export_NEU.bas, lager.json, MF_DACH_MAT.csv, PROJEKTINFO.md, Git-Repo.

---

## 🔴 KRITISCH

### K1: GitHub-Upload funktioniert nicht mehr – Git-Repo liegt im Lösch-Ordner
Das einzige Git-Repo (Remote `ddmbraun/Lagerverwaltung-FB300`) liegt unter:
`Prüfen bevor löschen\Backup\Lagerverw. FB300\.git`
Der aktive Projektordner `07_Lagerverwaltung-Excel\` hat **kein** Repo.

`NeueModule.GitHub_Export` berechnet `GIT_DIR` = Parent-Ordner der Mappe → das ist jetzt `D:\_KI-Projekte-2026\` (kein Repo) → `git add/commit/push` schlägt fehl (Fehler-Code im MsgBox). Letzter erfolgreicher Push: 07.06.2026.

**Zusatzgefahr:** Läge in `D:\_KI-Projekte-2026\` jemals ein Repo, würde `git add -A` **alle** Projekte (PAM, HV, …) committen und pushen.

**Lösung:**
1. Repo-Ordner `Lagerverw. FB300` aus „Prüfen bevor löschen" zurückholen ODER frisch klonen: `git clone https://github.com/ddmbraun/Lagerverwaltung-FB300.git` und die aktuellen Dateien (xlsm, .bas, lager.json) hineinlegen. Achtung: Im Backup-Repo liegen uncommittete Änderungen vom 07.06.
2. In `GitHub_Export` GIT_DIR absichern: vor dem `git add` prüfen `If Dir(GIT_DIR & ".git", vbDirectory) = "" Then MsgBox "Kein Git-Repo!": Exit Sub`.
3. Besser: GIT_DIR = `ThisWorkbook.Path` setzen (Mappe direkt ins Repo legen) statt Parent-Logik.

### K2: lager.json ist Datenmüll – Spalten-Mapping passt nicht zum Sheet
Beleg aus der exportierten lager.json (6.983 Einträge):
```json
{"nr":"#","artnr":"Art.-Nr.","artikel":"Artikel", ...}              ← Headerzeile als Artikel
{"nr":"Bohrer für Metall 5,0 mm ...","artnr":"5,5","artikel":"3","ean":"1,67","vk":"5", ...}
```
Artikelname steht unter `nr`, **VK unter `artnr`**, **Bestand unter `artikel`**, **EK unter `ean`**, die Excel-Zeilennummer unter `vk`. Spalten G–L liefern Altdaten.

Ursache: `ExportLagerJSON` erwartet Schnellansicht-Layout B=#, C=ArtNr, D=Artikel, E=EAN, F=VK, G=EK, H=Bestand … ab Zeile 3. `Schnellansicht_Aktualisieren` schreibt aber A=EAN, B=Artikel, C=VK, D=Anzahl, E=EK, F=Zeilenverweis ab Zeile 4.

**Dazu:** EK-Einkaufspreise landen so (im Feld `ean`) auf GitHub. **Prüfen, ob das Repo öffentlich ist!**

**Lösung:**
1. Export direkt aus dem **Artikel-Sheet** mit `Spalte_Finden()` statt fester Spaltennummern – dann ist es layoutunabhängig. Startzeile 5 (V3-Layout).
2. Zahlen als echte JSON-Zahlen mit Punkt exportieren (`Str(...)` + Trim statt deutschem Komma-String "1,67").
3. UTF-8 schreiben (ADODB.Stream statt `Print #`), sonst wird € zu `�`.
4. Entscheiden, ob EK überhaupt exportiert werden soll.

### K3: ArtikelFix.bas ist abgeschnitten
Die Datei endet in Zeile 286 mitten im Kommentar `'  5) Filter loeschen (`.
Es fehlen mindestens `Filter_Loeschen_Fix` und `ArtikelSheet()` – beide werden im Modul selbst **und** in den per `Events_Neu_Installieren` installierten Sheet-Events aufgerufen. Ein Re-Import dieser Datei erzeugt „Sub oder Function nicht definiert".

**Lösung:** Modul im VBA-Editor (Alt+F11) frisch exportieren und die Datei im Ordner ersetzen. **Diese Datei keinesfalls importieren oder auf GitHub als Stand sichern.** (Vermutlich Folge der bekannten Mount-/Kopier-Problematik – siehe CLAUDE.md-Warnungen der anderen Projekte.)

---

## 🟠 HOCH

### H1: Schnellansicht-Funktionen nutzen drei verschiedene Layouts
- `Schnellansicht_Aktualisieren` schreibt A–F (EAN, Artikel, VK, Anzahl, EK, Zeilenverweis)
- `Schnellansicht_Suchen`-Popup liest Bestand aus Spalte **5 (=EK!)** und Einheit aus Spalte **6 (=Zeilenverweis!)**
- `Schnellansicht_EK_Toggle` blendet Spalte **G** um (dort steht gar kein EK)
- Doppelklick-Handler (`Setup_Schnellansicht_DoubleClick`) liest EAN aus Spalte 5, Artikel aus Spalte 4

**Lösung:** EIN Layout verbindlich festlegen (am besten als Konstanten an einer Stelle) und alle vier Funktionen darauf umstellen. Steht als „Schnellansicht-Suche überprüfen" schon in PROJEKTINFO – das hier ist die konkrete Ursachenliste.

### H2: Artikel-Leseschleifen starten bei Zeile 3 statt 5
`Schnellansicht_Aktualisieren`, `SucheScanner_Aktualisieren`, `SchnellDetail_Laden`, `InvSuche_Suchen`, `Inventur_Befuellen` u. a. lesen `For i = 3 To lastRow`. Im V3-Layout ist Zeile 4 die Überschriftenzeile → der Header („ARTIKEL", „VK-PREIS €" …) wird als Artikel übernommen. Genau diese Geister-Einträge stehen am Anfang der lager.json.

**Lösung:** Startzeile zentral definieren (`Const ART_DATEN_START = 5`) und überall verwenden.

### H3: DuplikatPruefer findet im V3-Layout keine Spalten mehr
Sucht Header fest in **Zeile 2** (altes Layout). In V3 steht dort das Suchfeld → Abbruch mit „Spalte ARTIKEL nicht gefunden". Auch `anzArtikel = lastRow - 2` stimmt nicht mehr.

**Lösung:** `LagerMakros.Spalte_Finden()` verwenden, Daten ab Zeile 5.

---

## 🟡 MITTEL

### M1: GitHub_Export_NEU.bas ist gefährliche Altlast
Hartkodiertes `GIT_DIR = "...\Lagerverw. FB300\"` (existiert dort nicht mehr), exportiert nur 2 Module, kein „nichts zu committen"-Schutz. Laut PROJEKTINFO veraltet.
**Lösung:** Modul aus der Mappe entfernen (falls noch importiert) und Datei nach `Prüfen bevor löschen\` verschieben, damit niemand versehentlich die falsche `GitHub_Export`-Variante startet.

### M2: CSV-Import überschreibt gepflegte VK-Preise
Bei jedem Re-Import wird VK = Netto-EK × 1,35 neu gesetzt – manuell angepasste VK-Preise gehen verloren.
**Lösung:** VK nur setzen, wenn Zelle leer ist, oder vorher abfragen.
Weitere Punkte: `ZahlAusCSV` macht aus „1.234,56" eine 0 (Tausenderpunkt; aktuell in der CSV nicht vorhanden – geprüft – aber absichern: erst Punkte entfernen, dann Komma→Punkt). Neue Artikel werden ohne Bestände-Sheet-Eintrag angefügt (inkonsistent zu `NeuerArtikel_Speichern`).

### M3: Keine Duplikatprüfung beim Artikel-Anlegen
`NeuerArtikel_Speichern` fügt immer an, auch wenn ARTIKELNR/EAN schon existiert. `SucheScanner_ArtikelLaden` lädt zudem Bestandsartikel ins „Neuer Artikel"-Formular → Speichern erzeugt eine Dublette statt Update.
**Lösung:** Vor dem Anfügen per Dictionary auf ARTIKELNR prüfen; bei Treffer fragen „Aktualisieren statt neu anlegen?".

### M4: InvSuche-UserForm-Fehler (bekanntes offenes Problem)
`InvSuche_Form_Installieren` löscht und erzeugt die Form bei **jedem** Setup per `VBComponents.Add(3)`. Das Laufzeit-Erzeugen schreibt temporäre .frx-Dateien – typische Ursache für „Fehler beim Zugriff auf Pfad/Datei" (Virenscanner, OneDrive, Temp-Rechte).
**Lösung:** Form nur erzeugen, wenn sie fehlt (Existenz-Check statt Remove+Add). Oder ganz auf die UserForm verzichten und das bereits vorhandene `InvEingabe`-Sheet als Popup nutzen – gleiche Funktion, kein VBProject-Zugriff nötig.

### M5: LagerMakros.bas ohne `Option Explicit`
Einziges Modul ohne. Tippfehler in Variablennamen fallen nicht auf (z. B. funktioniert `vbA.UserForms.Add` in `InvSuche_ArtikelWaehlen` nur, weil VBA das zufällig als VBA-Bibliothek auflöst).
**Lösung:** `Option Explicit` ergänzen, dann Debuggen → Kompilieren und gemeldete Stellen fixen.

---

## 🟢 KLEINIGKEITEN / HINWEISE

- **ZuAbgang/Etikett:** Mindest-Zeilen-Check `g_LetzteZeile < 3` müsste in V3 `< 5` heißen (sonst sind Toolbar-/Headerzeile „buchbar"). Spaltenindizes werden ohne 0-Check benutzt → Laufzeitfehler 1004, falls eine Überschrift fehlt.
- **Etikett/ZPL:** Artikelnamen mit `^` oder `~` zerschießen das ZPL-Format – ggf. Zeichen ersetzen.
- **Export-Liste unvollständig:** `GitHub_Export` exportiert CSV_Import.bas nicht → die Datei im Ordner kann vom Stand in der Mappe abweichen.
- **PROJEKTINFO.md veraltet:** Pfade (`...\Lagerverw. FB300\Lagerverwaltung_2026_V3\...`) und Git-Root-Beschreibung entsprechen nicht mehr der realen Ordnerlage. Nach Repo-Umzug (K1) aktualisieren.
- **Backup-Lage dünn:** Nur ein xlsm-Backup vom 07.06. Vor den Korrekturen Backup im Windows-Explorer anlegen (nicht per Sandbox-Shell – bekannte Truncation-Gefahr).
- **Positiv:** lager.json ist syntaktisch gültiges JSON; CSV hat saubere 37-Felder-Struktur ohne Quoting-Fallen; die Suchlogik (Mehrwort-AND, Zahl→EAN/ArtNr) ist konsistent umgesetzt; `Spalte_Finden` mit Artikel-Sonderfall ab Zeile 3 ist ein guter zentraler Helfer.

---

## Empfohlene Reihenfolge

1. **K3** – ArtikelFix.bas frisch aus Excel exportieren (sonst geht ein kaputter Stand auf GitHub)
2. **K1** – Git-Repo zurückholen/neu klonen, GIT_DIR-Logik + Sicherheitscheck einbauen
3. **K2 + H2** – Export aus Artikel-Sheet neu schreiben (Spalte_Finden, ab Zeile 5, JSON-Zahlen, UTF-8); EK-Export klären, Repo-Sichtbarkeit prüfen
4. **H1** – Schnellansicht-Layout vereinheitlichen
5. **H3, M1–M5** nach Bedarf
