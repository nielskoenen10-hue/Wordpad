# Logboek

Een WordPad-achtige server-logboek app met foto's, automatisch bewaren en bestandsimport.

## Installeren

```bash
pip install flask Pillow python-docx
```

## Starten

```bash
python app.py
```

Open daarna `http://localhost:5000` in de browser.

## Functies

- Vrij typen met opmaakwerkbalk (vet, cursief, onderstrepen, lettertype, -grootte, kleur)
- Foto's invoegen via Ctrl+V, slepen, of de Foto-knop
- **.txt en .docx bestanden openen** via de Openen-knop (wordt een nieuwe notitie)
- Datum, tijd en naam automatisch bij elke notitie
- Automatisch bewaren tijdens het typen
- Notities vergrendelen voor betrouwbaarheid
- Zoeken door alle notities
- JSON-export voor back-up

## Data

- `logboek.db` — SQLite database (klein, alleen tekst)
- `fotos/` — gecomprimeerde afbeeldingen (max 1600px breed, JPEG)

Back-up = deze twee kopiëren.
