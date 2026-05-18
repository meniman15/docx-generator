# docx-generator

Fill Word (`.docx`) templates from JSON using `{{placeholder}}` tags.

## Contents

- `templateToDocx.py` — template engine (placeholders, collections, tables, HTML lists)
- `app.py` — Flask API (`POST /api/generate-docx`)
- `Dockerfile` — container image

Place your template file (e.g. `templateV2.docx`) next to `app.py` or set `TEMPLATE_PATH`.

## CLI

```bash
pip install -r requirements.txt
python templateToDocx.py template.docx data.json -o output.docx
```

## API

```bash
python app.py
# POST http://localhost:8080/api/generate-docx
# Body: JSON object with template field values
```

## Docker

```bash
docker build -t docx-generator .
docker run -p 8080:8080 docx-generator
```
