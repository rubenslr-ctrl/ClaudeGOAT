# Kleinanzeigen Häuser-Scraper

Sammelt Haus-Inserate (Kauf + Miete, deutschlandweit, private + gewerbliche Anbieter)
von kleinanzeigen.de und speichert **nur** Inserate, die Straße, Hausnummer, Baujahr und
Gebäudetyp selbst angeben, als Excel-Datei. Es werden keine Werte ergänzt oder geschätzt.

**Spalten:** Stadt | PLZ | Straße | Hausnummer | Gebäudetyp | Baujahr | Wohnfläche (m²) |
Sonstige Informationen (Preis, Heizung, Energieträger, Energieklasse, Verbrauch, …) | Quelle (Link)

> Der Scraper läuft auf deinem eigenen PC. In Cloud-Umgebungen/Rechenzentren wird
> kleinanzeigen.de häufig blockiert.

## Start in VS Code (empfohlen)

1. Ordner `ClaudeGOAT` in VS Code öffnen.
2. `Terminal` → `Task ausführen …` (bzw. `Strg+Shift+P` → „Tasks: Run Task“).
3. **Scraper: Test (10 Häuser)** wählen. Beim ersten Mal wird automatisch alles installiert
   (Python per winget, falls es fehlt, eine virtuelle Umgebung `.venv` und die Pakete).
4. Wenn der Test Treffer (`✔`) liefert: **Scraper: Starten (500 Häuser)**.

## Start ohne VS Code

| | Windows (Doppelklick) | macOS / Linux (Terminal) |
|---|---|---|
| Installieren | `installieren.bat` | `./installieren.sh` |
| Probelauf (10 Häuser) | `test_starten.bat` | `./starten.sh --test` |
| Richtiger Lauf (500) | `starten.bat` | `./starten.sh` |

## Optionen

```
--ziel 200        Anzahl vollständiger Treffer (Standard 500)
--nur-kauf        nur Häuser zum Kauf
--nur-miete       nur Häuser zur Miete
--datei name.xlsx anderer Name für die Excel-Datei
--test            Probelauf mit 10 Treffern in eigene Dateien
```

Beispiel: `starten.bat --ziel 200 --nur-kauf`

## Gut zu wissen

- **Dauer:** Zwischen den Abrufen wartet das Skript 2,5–5 Sekunden. 500 Treffer dauern
  mehrere Stunden – am besten über Nacht laufen lassen.
- **Abbrechen & fortsetzen:** `Strg+C` bricht ab. Beim nächsten Start geht es dort weiter
  (Stand in `fortschritt.json`). Für einen kompletten Neustart `fortschritt.json` löschen.
- **Excel nicht geöffnet lassen**, während das Skript läuft – sonst kann es nicht speichern.
- **Sperre / Captcha:** Das Skript pausiert und hört bei anhaltender Sperre auf. Einfach
  ein paar Stunden später erneut starten.
- **Keine Treffer trotz vieler Seiten?** Dann hat Kleinanzeigen wahrscheinlich den
  Seitenaufbau geändert und die Feldnamen im Skript müssen angepasst werden.
- **Datenschutz:** Adressen privater Anbieter sind personenbezogene Daten (DSGVO) – für
  Veröffentlichungen nur aggregiert bzw. auf PLZ-/Ortsebene verwenden.
