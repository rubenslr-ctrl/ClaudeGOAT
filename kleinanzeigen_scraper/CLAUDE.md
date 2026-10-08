# Übergabe: Kleinanzeigen Häuser-Scraper

Kontext aus einer früheren Claude-Code-Sitzung (Cloud). Mit dem Nutzer **auf Deutsch** kommunizieren.

## Ziel
Wissenschaftliches Paper: **500 Gebäude (nur Häuser: EFH, RFH, DHH, MFH …)** von kleinanzeigen.de,
deutschlandweit, Kauf + Miete, private + gewerbliche Anbieter, als **Excel** mit den Spalten:

Stadt | PLZ | Straße | Hausnummer | Gebäudetyp | Baujahr | Wohnfläche (m²) |
Sonstige Informationen (Preis, Heizungsart, Energieträger, Energieeffizienzklasse, Energiebedarf/-verbrauch, …) | Quelle (Link)

## Harte Anforderungen des Nutzers
- **Keine Daten erfinden oder schätzen.** Jeder Wert stammt aus dem Inserat, Quelle-Link in jeder Zeile.
- **Straße + Hausnummer sind Pflicht**, ebenso Baujahr und Gebäudetyp – sonst Inserat verwerfen.
- „Wert“ = Angebotspreis (bei Miete: Miete), kein Gutachterwert – so benennen.

## Stand
- Skript: `kleinanzeigen_haeuser.py` (requests + BeautifulSoup + openpyxl), Optionen `--test` (10 Treffer,
  eigene Dateien), `--ziel N`, `--nur-kauf`, `--nur-miete`, `--datei`. Start: `./starten.sh --test` bzw.
  `./starten.sh` (legt `.venv` an und installiert `requirements.txt`). VS-Code-Tasks in `../.vscode/tasks.json`.
- Ausgabe im Skriptordner: `test_haeuser_kleinanzeigen.xlsx` / `haeuser_kleinanzeigen.xlsx`, Fortschritt in
  `fortschritt*.json` (Fortsetzen nach Abbruch). Diese Dateien sind per `.gitignore` ausgeschlossen (DSGVO).
- Suche: `s-haus-kaufen/…/c208` und `s-haus-mieten/…/c205`, aufgeteilt in Preisbänder
  (`/preis:MIN:MAX/`, `/seite:N/`), weil eine Suche nur ~50 Seiten hat. Pause 2,5–5 s pro Abruf.

## Diagnose vom echten Kleinanzeigen (Mac des Nutzers, Okt. 2026)
- Suchseite: HTTP 200, **kein `article.aditem` mehr** (neues Tailwind-Markup: `ul > li > article`).
  Inserat-Links `/s-anzeige/<slug>/<id>-208-<ortid>`; ~54 eindeutige Links pro Seite. Preisfilter-URL funktioniert.
  → Skript wurde angepasst: Links per Regex, gefiltert nach Kategorienummer (208/205). **Live noch nicht bestätigt.**
- Inseratseite (getestet: Bauträger-Anzeige „allkauf“ ohne Adresse): `#viewad-title`, `#viewad-locality` (2×),
  `#viewad-price`, `li.addetailslist--detail` mit `.addetailslist--detail--value` (Wohnfläche, Zimmer, Haustyp,
  Etagen, Provision, Grundstücksfläche …), `.checktag` vorhanden. `#street-address` **nicht** vorhanden –
  bei dieser Anzeige erwartbar, aber **bei einem Inserat mit Adresse noch nicht geprüft**.
  Kein „Baujahr“-Feld in dieser Anzeige; Heizungsart stand nur im Beschreibungstext.

## Nächste Schritte
1. `./starten.sh --test` ausführen und Ausgabe prüfen (pro Seite: „X Inserate, Y neu geprüft, Z vollständig“).
2. Falls dauerhaft „0 vollständig“: ein Inserat **mit sichtbarer Adresse** abrufen und prüfen, wo Straße/Hausnr.
   und Baujahr im HTML stehen (Selektoren in `parse_inserat` anpassen). Feldnamen in `SONSTIGES_FELDER` prüfen.
3. Danach Volllauf (500), am besten über Nacht.
4. Vom Nutzer gewünschte/angebotene Erweiterungen (noch offen, vorher fragen):
   Kürzel-Spalte EFH/RFH/DHH/MFH, Abrufdatum, Herkunftsspalte (Feld vs. Beschreibung),
   Baujahr aus Fließtext nur wenn eindeutig und klar markiert.

## Grenzen
- **Keine IP-Rotation/Proxys oder sonstiges Umgehen von Sperren** – wurde bewusst abgelehnt. Bei 403/429/Captcha
  pausiert das Skript bzw. bricht ab. 500 Treffer sind mit normalem, langsamem Abruf machbar.
- Datenschutz: Adressen nicht committen/veröffentlichen; im Paper nur aggregiert bzw. auf PLZ-Ebene.

## Umgebung des Nutzers
- macOS, Python 3.9 (LibreSSL – urllib3-Warnung wird im Skript unterdrückt), Repo `~/ClaudeGOAT`.
- Lokales Repo steht auf `main` mit vielen eigenen, nicht committeten Änderungen (andere Projekte: data/, docs/, …).
  **Nicht den Branch wechseln.** Scraper-Dateien kamen per
  `git checkout origin/kleinanzeigen-scraper -- kleinanzeigen_scraper .vscode` in den Arbeitsordner.
- Alternative Datenquelle: Apify-MCP mit Actor `clearpath/kleinanzeigen-immobilien-api-pro`
  (~0,0015 $/Inserat) – Apify-Monatslimit des Nutzers war erreicht.
