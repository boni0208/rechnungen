# Rechnungs-App Hausbau Gerlafingen

Handy-Formular zum Erfassen von Rechnungen der Firma Boniface anhand der Mängelliste des Architekten.
Jede Rechnung wird als **PDF + Excel** (A4-Layout wie Rechnung 1023) erstellt, im GitHub-Repo abgelegt,
zum Download angeboten und auf Wunsch per E-Mail verschickt.

## Ablauf in der App
1. **Rechnung** – Nummer (wird automatisch hochgezählt), Rechnungsdatum, Ausführungszeitraum.
2. **Position wählen** – z.B. «273.3 · Innentüren». Die Einzelmängel aus der Liste erscheinen als Vorschlag zum Abhaken.
   Für Arbeiten ausserhalb der Liste: «Zusatzleistung».
3. **Leistungen erfassen** – pro Zeile Text, Menge, Einheit (Stk., Std., pauschal …), Ansatz, Bezug zum Mängelpunkt
   und optional Unterpunkte (erscheinen auf der Rechnung als einzelne «–»-Zeilen).
   Knöpfe für schnelle Zeilen: Material, Arbeit, Lieferung, Pauschale, Fahrspesen.
4. **Rechnung erstellen** – Total prüfen, erstellen, PDF/Excel herunterladen, optional per Mail senden.

Auf der Rechnung erhält jede Position eine eigene Nummer der Firma (1, 1.1, 1.2 …) mit Hinweis auf die
Mängelliste (Pos. 273.3, Mängelpunkt 273.3-01). Ohne MWST-Nr. wird keine MWST ausgewiesen; mit MWST-Nr. 8.1 %.

## Einrichten (einmalig, ca. 20 Minuten)

### 1. GitHub-Repo
1. Auf github.com ein neues **privates** Repo anlegen, z.B. `hausbau-rechnungen`.
2. Alle Dateien aus diesem Ordner hochladen (*Add file → Upload files*). **Nicht** hochladen: eine echte `secrets.toml`.

### 2. GitHub-Token (damit die App Rechnungen speichern darf)
*Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token*
- Repository access: **Only select repositories** → `hausbau-rechnungen`
- Permissions → Repository → **Contents: Read and write**
- Token kopieren.

Die Rechnungen landen im Branch `rechnungen` (wird automatisch angelegt), Ordner `rechnungen/<Jahr>/`,
plus `rechnungen/journal.csv` als Liste aller Rechnungen. Getrennter Branch = die App startet beim Speichern nicht neu.

### 3. App veröffentlichen (Streamlit Community Cloud, gratis)
1. Auf share.streamlit.io mit GitHub anmelden → **Create app** → Repo `hausbau-rechnungen`, Branch `main`, Datei `app.py`.
2. **Advanced settings → Secrets**: Inhalt von `.streamlit/secrets.toml.example` einfügen und die Werte anpassen
   (Passwort, Firmendaten, Token, Repo-Name, optional Mail).
3. Deploy. Den Link auf dem Handy deines Vaters öffnen → *Teilen → Zum Home-Bildschirm*.

### 4. E-Mail (optional)
Für Gmail: Google-Konto → Sicherheit → 2-Faktor aktivieren → **App-Passwörter** → neues Passwort erzeugen
und im Abschnitt `[mail]` eintragen. Abschnitt weglassen = kein Mailversand.

## Lokal testen
```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # Werte anpassen
streamlit run app.py
```

## Dateien
| Datei | Zweck |
|---|---|
| `app.py` | Formular (Streamlit) |
| `invoice.py` | Rechnungsmodell, Excel- und PDF-Ausgabe |
| `storage.py` | Ablage im GitHub-Repo + Journal |
| `mailer.py` | Versand per SMTP |
| `data/maengelliste.json` | Positionen und Einzelmängel der Mängelliste (bei neuer Liste hier anpassen) |

## Später erweiterbar
- AI-Strukturierung von Freitext (z.B. mit einem LLM-API-Key) – bewusst noch nicht eingebaut.
- Automatisches Nachführen der Übersichts-Excel aus `journal.csv`.
