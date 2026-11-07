"""Сборка сайта диагностики (EduApp demo).

Запуск:  python build.py data/demo_students.json            (демо-данные, см. scripts/generate_demo_data.py)
        python build.py results.docx "English Diagnostic Grade 4.xlsx" [...]   (реальные выгрузки)
Результат: index.html рядом с этим файлом (template.html + данные).

.docx — таблица с разбором по 5 навыкам; .xlsx — выгрузка участников
(только общий балл XP, без навыков; класс берётся из «Grade N» в имени файла).

Телефоны в index.html открытым текстом не попадают. Имена есть (для подсказок при вводе),
результат открывается по имени ребёнка. Если одинаковое имя у разных детей,
результат открывается по имени + последним 4 цифрам телефона.
Функции norm/cyrb53/keystream должны совпадать с JS в template.html.
"""
import json
import os
import shutil
import random
import re
import sys
from pathlib import Path


HERE = Path(__file__).parent
OUT = HERE / "dist"  # эта папка публикуется (Vercel / любой статический хостинг)
SITE = os.environ.get("SITE_URL", "https://diagnostic.eduapp.example.com")  # адрес сайта для og-тегов
M = 0xFFFFFFFF

TR = {
    "а": "a", "ә": "a", "б": "b", "в": "v", "г": "g", "ғ": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "i", "і": "i", "к": "k", "қ": "k", "л": "l", "м": "m",
    "н": "n", "ң": "n", "о": "o", "ө": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ұ": "u", "ү": "u", "ф": "f", "х": "h", "һ": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
}


def imul(a, b):
    return (a * b) & M


def cyrb53(s, seed=0):
    h1 = 0xDEADBEEF ^ seed
    h2 = 0x41C6CE57 ^ seed
    for ch in s:
        c = ord(ch)
        h1 = imul(h1 ^ c, 2654435761)
        h2 = imul(h2 ^ c, 1597334677)
    h1 = imul(h1 ^ (h1 >> 16), 2246822507)
    h1 ^= imul(h2 ^ (h2 >> 13), 3266489909)
    h2 = imul(h2 ^ (h2 >> 16), 2246822507)
    h2 ^= imul(h1 ^ (h1 >> 13), 3266489909)
    return 4294967296 * (2097151 & h2) + (h1 & M)


def words(name):
    s = "".join(TR.get(ch, ch) for ch in name.lower())
    s = s.replace("zh", "j").replace("kh", "h").replace("y", "i").replace("w", "u").replace("q", "k")
    s = re.sub(r"[^a-z]+", " ", s)
    s = re.sub(r"(.)\1+", r"\1", s)
    return sorted(w for w in s.split() if w)


def b36(n):
    out = ""
    while True:
        n, r = divmod(n, 36)
        out = "0123456789abcdefghijklmnopqrstuvwxyz"[r] + out
        if n == 0:
            return out


def key_id(key):
    return b36(cyrb53(key, 1))


def seal(text, key):
    data = text.encode("utf-8")
    return bytes(b ^ (cyrb53(f"{key}:{i}", 99) & 255) for i, b in enumerate(data)).hex()


def num(v):
    v = v.strip().replace(",", ".")
    return round(float(v), 1) if v else 0.0


def read_docx(path):
    import docx  # python-docx, нужен только для реальных выгрузок
    d = docx.Document(path)
    grade = ""
    for p in d.paragraphs:
        m = re.search(r"(\d+)\s*(?:st|nd|rd|th)?\s*grade", p.text, re.I)
        if m:
            grade = m.group(1)
            break
    rows = []
    for r in d.tables[0].rows[1:]:
        c = [x.text.strip() for x in r.cells]
        if not c[1]:
            continue
        phones = []
        for cell in (c[9], c[8]):
            for part in re.split(r"[;,/]", cell):
                digits = re.sub(r"\D", "", part)
                if len(digits) >= 4 and digits[-4:] not in phones:
                    phones.append(digits[-4:])
        rows.append({
            "name": c[1], "grade": grade, "cls": grade, "phones": phones, "done": True,
            "s": [num(c[2]), num(c[3]), num(c[4]), num(c[5]), num(c[6]), num(c[7])],
        })
    return rows


def last4(cell):
    out = []
    for part in re.split(r"[;,/]", str(cell or "")):
        digits = re.sub(r"\D", "", part)
        if len(digits) >= 4 and digits[-4:] not in out:
            out.append(digits[-4:])
    return out


def read_xlsx(path):
    """Выгрузка участников: ФИО, Класс, Статус, XP, телефоны. Навыков нет — только общий балл."""
    import openpyxl
    grade = re.search(r"Grade\s*(\d+)", Path(path).name, re.I).group(1)
    ws = openpyxl.load_workbook(path, data_only=True).active
    head = [str(h).strip() for h in next(ws.iter_rows(max_row=1, values_only=True))]
    col = {h: i for i, h in enumerate(head)}
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        name = str(r[col["ФИО"]] or "").strip()
        if not name or name.lower().startswith("admin"):
            continue
        done = r[col["Статус"]] == "Все этапы пройдены"
        cls = re.match(r"\s*(\d+)", str(r[col["Класс"]] or ""))
        rows.append({
            "name": name, "grade": grade, "cls": cls.group(1) if cls else grade,
            "phones": last4(r[col["Телефон ученика"]]) + last4(r[col["Телефоны родителей"]]),
            "done": done, "s": [float(r[col["XP"]] or 0)] if done else [-1],
        })
    return rows


def read_json(path):
    """Демо-данные: список учеников в том же формате, что дают read_docx / read_xlsx."""
    return json.loads(Path(path).read_text(encoding="utf-8"))


def merge_same_child(students):
    """Один ребёнок прошёл тест несколько раз (одно имя + общий телефон) — оставляем одну запись:
    тест своего класса, иначе пройденный с лучшим баллом."""
    groups = {}
    for s in students:
        groups.setdefault(" ".join(words(s["name"])), []).append(s)
    out = []
    for group in groups.values():
        clusters = []
        for s in group:
            for c in clusters:
                if not s["phones"] or set(s["phones"]) & {p for x in c for p in x["phones"]}:
                    c.append(s)
                    break
            else:
                clusters.append([s])
        for c in clusters:
            c.sort(key=lambda x: (x["done"], x["grade"] == x["cls"], len(x["s"]) > 1, x["s"][0]), reverse=True)
            out.append(c[0])
    return out


def main(paths):
    readers = {".xlsx": read_xlsx, ".json": read_json}
    students = [s for p in paths for s in readers.get(Path(p).suffix.lower(), read_docx)(p)]
    students = merge_same_child(students)
    random.seed(7)
    random.shuffle(students)

    scores = [[s["grade"]] + s["s"] for s in students]
    keys = {}
    ambiguous = []
    need_pin = []
    # Ключи: полное имя + каждое слово имени (если оно есть только у одного ученика)
    candidates = {}
    for i, s in enumerate(students):
        ws = words(s["name"])
        for v in {" ".join(ws)} | set(ws):
            candidates.setdefault(v, set()).add(i)
    full_names = {" ".join(words(s["name"])) for s in students}
    for key, idx in candidates.items():
        if len(idx) == 1:
            i = idx.pop()
            keys[key_id(key)] = [i, seal(students[i]["name"], key)]
        elif key in full_names:
            # Одинаковое полное имя у разных детей — различаем по последним 4 цифрам телефона
            need_pin.append(key_id(key))
            for i in idx:
                for pin in students[i]["phones"]:
                    k = f"{key}#{pin}"
                    keys[key_id(k)] = [i, seal(students[i]["name"], k)]
        else:
            ambiguous.append(key_id(key))  # одно слово у нескольких — просим полное имя

    data = {"s": scores, "k": keys, "a": sorted(ambiguous), "p": sorted(need_pin),
            "n": [s["name"] for s in students]}  # n — для подсказок
    html = (HERE / "template.html").read_text(encoding="utf-8")
    html = html.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("__SITE__", SITE)
    (HERE / "index.html").write_text(html, encoding="utf-8")  # готовая демо-страница в корне репозитория
    (OUT / "insta").mkdir(parents=True, exist_ok=True)
    (OUT / "index.html").write_text(html, encoding="utf-8")
    insta = (HERE / "insta.html").read_text(encoding="utf-8").replace("__SITE__", SITE)
    for d in ("insta", "ig", "instagram"):  # /instagram — актуальный адрес; WhatsApp кэширует старые превью
        (OUT / d).mkdir(exist_ok=True)
        (OUT / d / "index.html").write_text(insta, encoding="utf-8")
    shutil.copy(HERE / "assets" / "logo.png", OUT / "logo.png")
    shutil.copy(HERE / "vercel.json", OUT / "vercel.json")
    shutil.copy(HERE / "assets" / "favicon.svg", OUT / "favicon.svg")
    print(f"OK: {len(students)} оқушы, {len(keys)} кілт -> {OUT.name}/")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("python build.py <файл.json|файл.docx|файл.xlsx> [...]")
    main(sys.argv[1:])
