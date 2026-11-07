"""Generate a random demo dataset of diagnostic-test participants.

Output: data/demo_students.json — the same record format that build.py's
read_docx / read_xlsx produce, so `python build.py data/demo_students.json`
builds a fully working demo site. Every name, phone digit and score is random.

    python scripts/generate_demo_data.py [--count 160] [--seed 42]
"""
import argparse
import json
import random
from pathlib import Path

GIRLS = ["Айгерім", "Аружан", "Әсел", "Дана", "Жансая", "Мадина", "Томирис", "Інжу", "Аяулым", "Ақбота",
         "Камила", "Сабина", "Сымбат", "Айша", "Зере", "Медина", "Амина", "Еркежан", "Балнұр", "Гүлназ"]
BOYS = ["Ерлан", "Нұрсұлтан", "Әлихан", "Тимур", "Арсен", "Данияр", "Ернұр", "Нұрәли", "Елдос", "Айдар",
        "Санжар", "Дәулет", "Ислам", "Рамазан", "Нұрислам", "Әмір", "Алихан", "Ақжол", "Ильяс", "Самат"]
SURNAMES = ["Сапаров", "Нұрланов", "Ахметов", "Серікбаев", "Жұмабаев", "Қасымов", "Омаров", "Бекенов",
            "Тұрсынов", "Мұратов", "Әбенов", "Исаев", "Құдайбергенов", "Сейітов", "Байжанов", "Жақсылықов",
            "Оспанов", "Төлеуов", "Ермеков", "Қайратов"]


def surname(base, girl):
    return base[:-2] + "ова" if girl and base.endswith("ов") else base + "а" if girl and base.endswith("ев") else base


def clamp(v):
    return round(max(0.0, min(100.0, v)), 1)


def student(rng, grade):
    girl = rng.random() < 0.5
    name = f"{surname(rng.choice(SURNAMES), girl)} {rng.choice(GIRLS if girl else BOYS)}"
    phones = ["".join(rng.choice("0123456789") for _ in range(4)) for _ in range(rng.choice((1, 1, 2)))]
    cls = grade if rng.random() > 0.08 else str(int(grade) + rng.choice((-1, 1)))  # some kids take another grade's test
    ability = rng.gauss(68 - (int(grade) - 3) * 2, 14)  # harder tests in higher grades
    if rng.random() < 0.07:  # started but did not finish
        return {"name": name, "grade": grade, "cls": cls, "phones": phones, "done": False, "s": [-1]}
    if rng.random() < 0.3:  # export with the total score only (no skill breakdown)
        return {"name": name, "grade": grade, "cls": cls, "phones": phones, "done": True, "s": [clamp(ability)]}
    # Writing, Vocabulary, Reading, Grammar, Listening — writing is usually the weakest skill
    bias = [-9, 4, 1, -3, 5]
    skills = [clamp(ability + b + rng.gauss(0, 9)) for b in bias]
    total = clamp(sum(skills) / len(skills))
    return {"name": name, "grade": grade, "cls": cls, "phones": phones, "done": True, "s": [total, *skills]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--count", type=int, default=160)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--out", default=str(Path(__file__).resolve().parent.parent / "data" / "demo_students.json"))
    a = ap.parse_args()
    rng = random.Random(a.seed)
    rows = [student(rng, rng.choice("3456")) for _ in range(a.count)]
    # a guaranteed name clash so the "last 4 phone digits" path can be tried in the demo
    rows.append({"name": "Сапарова Айгерім", "grade": "4", "cls": "4", "phones": ["1111"], "done": True, "s": [82.4, 70.1, 88.0, 85.5, 79.9, 88.5]})
    rows.append({"name": "Сапарова Айгерім", "grade": "5", "cls": "5", "phones": ["2222"], "done": True, "s": [64.2, 51.0, 70.3, 66.8, 60.4, 72.5]})
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(rows)} demo students -> {a.out}")


if __name__ == "__main__":
    main()
