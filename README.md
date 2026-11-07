# English Diagnostic Results

A tiny static site where parents type their child's name and instantly see the result of an English placement test: overall score, five skills (writing, vocabulary, reading, grammar, listening) compared with the grade average, strengths and focus areas — in Kazakh and Russian. No backend: the data is baked into one HTML file, and names are never stored in plain text.

## Features
- **Lookup by name** with typo-tolerant matching (Cyrillic/Kazakh → Latin normalisation, word order does not matter, a surname or first name alone works when it is unique).
- **Privacy by design**: results are keyed by a `cyrb53` hash of the normalised name; the display name is XOR-sealed with a keystream derived from the same key, so the page cannot be scraped into a list of children with scores. Phone numbers never reach the page — only the last 4 digits are used, and only to tell apart children with identical names.
- **Ambiguity handling**: if one word matches several pupils the parent is asked for the full name; identical full names ask for the last 4 phone digits.
- Result card: animated score ring, level badge, per-skill bars with the grade average, strengths / what to practise, print-friendly layout.
- Pupils who started but did not finish the test get a "not finished yet" state; exports that only contain a total score render without the skill breakdown.
- KK / RU switch (remembered in `localStorage`), Open Graph tags, an `/instagram` redirect page with link-preview metadata.

## Tech stack
- Python 3 build script (`build.py`) — reads test exports (`.docx` table, `.xlsx` participant export, or demo `.json`), merges duplicate attempts, builds the hashed lookup index and injects it into the template.
- Vanilla HTML/CSS/JS (`template.html`) — no framework, no bundler; the JS hash functions mirror the Python ones byte-for-byte.
- Static hosting (Vercel config included, any static host works).

## Project structure
```
build.py                    # export(s) -> index.html + dist/ (deployable folder)
template.html               # page template, data placeholder /*__DATA__*/null
insta.html                  # /instagram redirect page with OG tags
index.html                  # prebuilt demo page (from the generated data)
data/demo_students.json     # randomly generated pupils
scripts/generate_demo_data.py
assets/                     # neutral placeholder logo / favicon
vercel.json
```

## How to run (demo mode)
Just open `index.html` in a browser (or `python -m http.server` and go to http://localhost:8000). Try a name from `data/demo_students.json`, e.g. `Нұрланова Еркежан`, or `Сапарова Айгерім` (two pupils share it — enter `2222` as the last 4 digits).

Rebuild from fresh random data:
```bash
python scripts/generate_demo_data.py --count 160 --seed 42
python build.py data/demo_students.json     # writes index.html and dist/
```
With real exports: `pip install -r requirements.txt`, then `python build.py results.docx "English Diagnostic Grade 4.xlsx"`. Set `SITE_URL` for the Open Graph links.

> All data in this repository is randomly generated; brand names are anonymized.

## Кратко (RU)

Статический сайт: родитель вводит имя ребёнка и сразу видит результат диагностики по английскому.
Общий балл и 5 навыков в сравнении со средним по классу; на казахском и русском.
Имена не хранятся в открытом виде: поиск по хэшу, используются только последние 4 цифры телефона.
Демо: откройте `index.html`; данные сгенерированы случайно через `scripts/generate_demo_data.py`.
