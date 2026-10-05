# Anmeldeseite – Einrichtung

Zwei Dateien:

- **index.html** – die Webseite, die du auf GitHub Pages hochlädst.
- **AppsScript.gs** – das Backend, das in einem Google Sheet läuft. Dieses
  Sheet ist gleichzeitig deine automatisch befüllte Excel-Liste.

## Warum ein Google Sheet als Backend?

Du wolltest: alle Anmeldungen landen automatisch in einer Excel-artigen
Liste, und wenn du dort manuell etwas rein- oder herausschreibst, soll
sich das auf der Webseite spiegeln. Ein normales Excel-File kann das
nicht (es liegt nicht im Netz), ein Google Sheet schon — und es lässt
sich jederzeit per "Datei → Herunterladen → Microsoft Excel (.xlsx)"
als echte Excel-Datei exportieren. Deshalb ist das Sheet selbst die
Datenbank: Jede Zelle entspricht genau einem Platz in der Tabelle auf
der Webseite (Zeile = Zeitslot, Spalte = Aktivität). Name in der Zelle
= besetzt, Zelle leer = frei — in beide Richtungen.

## 1. Google Sheet + Apps Script einrichten (ca. 5 Minuten)

1. Neues Google Sheet erstellen: https://sheets.new
2. Oben im Menü: **Erweiterungen → Apps Script**.
3. Den kompletten Inhalt von `AppsScript.gs` dort einfügen (vorhandenen
   Beispielcode ersetzen) und speichern.
4. Oben rechts auf **Bereitstellen → Neue Bereitstellung**.
   - Typ auswählen: **Web-App**
   - "Ausführen als": **Ich** (deine Google-Adresse)
   - "Wer hat Zugriff": **Jeder** (sonst müssten alle ein Google-Konto
     haben, um sich anzumelden)
   - Bereitstellen → Zugriff bestätigen (Google zeigt eine Warnung, weil
     es dein eigenes, unverifiziertes Script ist — "Erweitert" →
     "Trotzdem zu [Projektname] (unsicher)" wählen)
5. Du bekommst eine **Web-App-URL** (endet auf `/exec`). Diese URL in
   `index.html` eintragen, im Konfigurations-Block bei:
   ```js
   webAppUrl: "DEINE_APPS_SCRIPT_URL"
   ```

Jedes Mal, wenn du den Code in `AppsScript.gs` änderst, musst du erneut
"Bereitstellen → Neue Bereitstellung" ausführen (oder bei einer
bestehenden Bereitstellung auf "Verwalten" → Stift-Symbol → neue Version
wählen), sonst bleibt die alte Version aktiv.

## 2. Aktivitäten, Zeitslots und Limit einstellen

Alles an einer Stelle, ganz oben im `<script>`-Block von `index.html`:

```js
const CONFIG = {
  activities: ["Palatschinken", "Kassa", "Grill"],
  slots: ["Samstag 10:00", "Samstag 11:00", "Samstag 12:00", "Samstag 13:00"],
  maxTotalPerPerson: 3,
  webAppUrl: "DEINE_APPS_SCRIPT_URL"
};
```

- `activities`: eine Zeile pro Spalte/"Sache".
- `slots`: eine Zeile pro Zeitslot.
- `maxTotalPerPerson`: wie oft sich jemand insgesamt eintragen darf.

Mehr brauchst du nicht anzufassen. Das Google Sheet legt die passenden
Zeilen/Spalten beim ersten Zugriff automatisch an — du musst es nicht
vorher manuell einrichten.

Fügst du später eine neue Aktivität oder einen neuen Zeitslot hinzu,
trag sie hier ein; die passende Spalte/Zeile entsteht im Sheet
automatisch, sobald sich jemand dort einträgt.

## 3. Die eingebauten Regeln

- **Nicht zwei Sachen gleichzeitig**: Pro Zeitslot (Zeile) kann ein Name
  nur in einer Spalte stehen. Ein zweiter Versuch in derselben Zeile wird
  abgelehnt ("du bist schon angemeldet").
- **Nicht zu oft**: Wie oft ein Name insgesamt im Sheet vorkommen darf,
  bestimmt `maxTotalPerPerson`.
- **Vor- und Nachname**: Beides muss ausgefüllt sein; Namen werden ohne
  Rücksicht auf Groß-/Kleinschreibung und doppelte Leerzeichen verglichen,
  damit "Max Mustermann" und "max  mustermann" als dieselbe Person
  gelten.
- **Manuelle Bearbeitung im Sheet**: Schreibst du direkt im Sheet einen
  Namen in eine Zelle, ist der Platz ab der nächsten Aktualisierung (max.
  7 Sekunden, oder sofort per Klick auf "Jetzt aktualisieren") auf der
  Webseite belegt. Löschst du den Zellinhalt, ist er wieder frei.

Ohne Login kann es zwei echte Personen mit demselben Namen geben, die
dann vom System als eine Person behandelt würden — bei einem Freundes-/
Familien-/Vereinskreis in der Praxis meist kein Problem.

## 4. Auf GitHub Pages veröffentlichen

1. Neues Repository auf GitHub erstellen (z. B. `anmeldung`).
2. `index.html` dort hochladen (nur diese Datei braucht GitHub Pages;
   `AppsScript.gs` liegt ja schon in Google Apps Script).
3. **Settings → Pages** → als Branch `main` (Ordner `/root`) wählen →
   Speichern.
4. Nach ein bis zwei Minuten ist die Seite unter
   `https://DEIN-GITHUB-NAME.github.io/anmeldung/` erreichbar.

## Deine Excel-Liste

Das Google Sheet, das du in Schritt 1 angelegt hast, füllt sich live mit
jeder Anmeldung — kein Export-Schritt nötig. Wann immer du eine echte
.xlsx-Datei brauchst: im Sheet **Datei → Herunterladen → Microsoft
Excel (.xlsx)**.
