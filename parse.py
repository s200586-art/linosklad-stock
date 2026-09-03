#!/usr/bin/env python3
"""Собирает каталог ЛИНОСКЛАД из Excel 1С «Остатки линолеума по сериям для бота».

Остаток = сумма трёх тюменских складов (м² → пог. м).
Одна карточка = дизайн, ширины = варианты (разные ЦБ-артикулы).
"""
from __future__ import annotations

import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date
from pathlib import Path

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
TYUMEN_TITLES = ("тюмень",)
PRICE_FILE = Path(os.environ.get("PRICE_FILE", "/workspace/src/data/price-by-article.json"))
OUT_TS = Path(os.environ.get("OUT_TS", "/workspace/src/data/linoleum-stock.ts"))
OUT_JSON = Path(os.environ.get("OUT_JSON", "/workspace/src/data/stock-meta.json"))
OUT_LIVE = Path(os.environ.get("OUT_LIVE", "/workspace/public/stock-live.json"))

BRANDS = [
    ("tarkett", "Tarkett"),
    ("juteks", "Juteks"),
    ("juteкс", "Juteks"),
    ("forbo", "Forbo"),
    ("gerflor", "Gerflor"),
    ("sinteros", "Синтерос"),
    ("синтерос", "Синтерос"),
    ("комитекс", "Комитекс Лин"),
    ("komitex", "Комитекс Лин"),
    ("ivc", "IVC"),
    ("ideal", "Ideal"),
    ("grabo", "Grabo"),
    ("beauflor", "Beauflor"),
    ("magnatex", "Magnatex"),
    ("moduleo", "Moduleo"),
]

COLLECTION_BRAND = {
    "ultra": "Juteks",
    "master": "Juteks",
    "motive": "Juteks",
    "fortuna": "Juteks",
    "bazis": "Juteks",
    "imperia": "Juteks",
    "glory": "Juteks",
    "strong": "Juteks",
    "olympia": "Juteks",
    "premium": "Juteks",
    "stars": "Juteks",
    "story": "Juteks",
    "record": "Juteks",
    "evrika": "Juteks",
    "senator": "Juteks",
    "stimul": "Juteks",
    "flex": "Juteks",
    "pacific": "Juteks",
    "petergof": "Juteks",
}

TEXTURES = [
    "oak-light",
    "oak-honey",
    "walnut",
    "herringbone",
    "concrete",
    "tile-beige",
    "stone-gray",
    "marble",
    "chip-gray",
    "chip-blue",
    "marmoleum",
    "sport-green",
    "antislip",
]

RU = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "zh",
    "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o",
    "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "ts",
    "ч": "ch", "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu",
    "я": "ya",
}


def slugify(text: str) -> str:
    s = text.lower()
    s = "".join(RU.get(ch, ch) for ch in s)
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s[:80] or "linoleum"


def num(x) -> float:
    if x is None or x == "":
        return 0.0
    try:
        return float(str(x).replace("\xa0", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return 0.0


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def load_sheet(path: Path):
    with zipfile.ZipFile(path) as z:
        ss = []
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", NS):
            ss.append("".join(t.text or "" for t in si.findall(".//m:t", NS)))

        def val(c):
            t = c.attrib.get("t")
            v = c.find("m:v", NS)
            if v is None or v.text is None:
                isel = c.find("m:is", NS)
                if isel is not None:
                    return "".join(x.text or "" for x in isel.findall(".//m:t", NS))
                return None
            if t == "s":
                return ss[int(v.text)]
            return v.text

        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        rows: dict[int, dict[str, str | None]] = {}
        for row in sheet.findall("m:sheetData/m:row", NS):
            rnum = int(row.attrib["r"])
            cells: dict[str, str | None] = {}
            for c in row.findall("m:c", NS):
                ref = c.attrib["r"]
                col = re.match(r"[A-Z]+", ref).group(0)
                cells[col] = val(c)
            rows[rnum] = cells
        return rows


def parse_name(name: str) -> tuple[str, float | None, bool]:
    leftover = bool(re.search(r"остаток", name, re.I))
    n = re.sub(r"\s*\(остаток[^)]*\)", "", name, flags=re.I)
    n = re.sub(r"\s+", " ", n).strip()
    width = None
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*м\s*$", n, re.I)
    if not m:
        m = re.search(r"[-–]\s*(\d+(?:[.,]\d+)?)\s*м\s*$", n, re.I)
    if not m:
        m = re.search(r"\((\d+(?:[.,]\d+)?)\s*м\s*/", n, re.I)
    if not m:
        m = re.search(r"(\d+(?:[.,]\d+)?)\s*м", n, re.I)
    if m:
        width = float(m.group(1).replace(",", "."))
    n = re.sub(r"\s*\(\s*\d+[.,]?\d*\s*м\s*/[^)]*\)", " ", n)
    n = re.sub(r"\s+\d+(?:[.,]\d+)?\s*м\b", " ", n, flags=re.I)
    n = re.sub(r"_[A-Z]{0,4}\d{2,}\s*-?\s*$", "", n).strip(" -")
    n = re.sub(r"\s+", " ", n).strip()
    return n, width, leftover


def brand_of(name: str) -> tuple[str, str, str]:
    raw = re.sub(r"^линолеум\s+", "", name, flags=re.I)
    raw = re.sub(r"^рулонный винил\s+", "", raw, flags=re.I)
    raw = re.sub(r"^плитка пвх в рулоне\s+", "", raw, flags=re.I)
    tokens = raw.split()
    low = name.lower()
    brand = ""
    for key, label in BRANDS:
        if key in low:
            brand = label
            break
    if not brand and tokens:
        brand = COLLECTION_BRAND.get(tokens[0].lower(), tokens[0].title())
    rest = tokens[1:] if tokens and tokens[0].lower() in {brand.lower(), *(k for k, _ in BRANDS), *COLLECTION_BRAND} else tokens
    if brand.lower() == (tokens[0].lower() if tokens else ""):
        rest = tokens[1:]
    collection = rest[0] if rest else brand
    design = " ".join(rest[1:]) if len(rest) > 1 else (rest[0] if rest else raw)
    return brand, collection, design


def texture_of(name: str) -> str:
    n = name.lower()
    if "ёлоч" in n or "herring" in n:
        return "herringbone"
    if "мрам" in n or "marble" in n:
        return "marble"
    if "плит" in n or "tile" in n:
        return "tile-beige"
    if "бетон" in n or "concrete" in n:
        return "concrete"
    if "орех" in n or "walnut" in n:
        return "walnut"
    if "чип" in n or "dots" in n:
        return "chip-gray"
    if "мармол" in n:
        return "marmoleum"
    if "дуб" in n or "oak" in n:
        return "oak-honey" if any(x in n for x in ("honey", "медов", "gold")) else "oak-light"
    h = sum(map(ord, n)) % len(TEXTURES)
    return TEXTURES[h]


def category_of(name: str) -> str:
    n = name.lower()
    if "мармол" in n or "marmoleum" in n:
        return "marmoleum"
    if "гомоген" in n or "iq " in n or " sphera" in n:
        return "kommercheskiy-gomogen"
    if "коммерч" in n or "acczent" in n or "taralay" in n:
        return "kommercheskiy-geterogen"
    if any(x in n for x in ("supreme", "force", "sprint", "полукоммер")):
        return "polukommercheskiy"
    if any(x in n for x in ("5м", " 5 м")):
        return "kommercheskiy-geterogen"
    return "bytovoy"


def js_str(s: str) -> str:
    return json.dumps(s, ensure_ascii=False)


def main() -> int:
    src = Path(sys.argv[1] if len(sys.argv) > 1 else "/workspace/artifacts/ostatki-lino-2026-08-28.xlsx")
    if not src.exists():
        print(f"no file {src}", file=sys.stderr)
        return 1
    rows = load_sheet(src)
    header = rows.get(8, {})
    tyumen_cols = [
        col
        for col, title in header.items()
        if title and "тюмень" in str(title).lower()
    ]
    if not tyumen_cols:
        print("no tyumen columns", file=sys.stderr)
        return 1

    prices: dict[str, float] = {}
    if PRICE_FILE.exists():
        try:
            prices = {str(k): float(v) for k, v in json.loads(PRICE_FILE.read_text()).items()}
        except Exception:
            prices = {}

    items = []
    cur = None
    for r in range(10, max(rows) + 1):
        cells = rows.get(r, {})
        art = str(cells.get("A") or "").strip()
        name = str(cells.get("D") or "").strip()
        if not art or art.lower() == "итого":
            continue
        tyu = sum(num(cells.get(c)) for c in tyumen_cols)
        if name:
            base, width, leftover = parse_name(name)
            cur = {
                "article": art,
                "name": name,
                "base": base,
                "width": width,
                "leftover": leftover,
                "tyumen_m2": tyu,
            }
            items.append(cur)
        elif cur:
            cur["tyumen_m2"] = max(cur["tyumen_m2"], tyu) if False else cur["tyumen_m2"]

    groups: dict[str, list] = defaultdict(list)
    for it in items:
        if it["tyumen_m2"] <= 0 or not it["width"]:
            continue
        groups[it["base"]].append(it)

    products = []
    used_slugs: set[str] = set()
    for base, skus in sorted(groups.items(), key=lambda kv: kv[0].lower()):
        brand, collection, design = brand_of(base)
        by_w: dict[float, dict] = {}
        leftover = False
        for sku in skus:
            w = float(sku["width"])
            lin = round(sku["tyumen_m2"] / w, 2)
            if lin <= 0:
                continue
            prev = by_w.get(w)
            if prev:
                prev["linear"] = round(prev["linear"] + lin, 2)
                prev["m2"] = round(prev["m2"] + sku["tyumen_m2"], 3)
            else:
                by_w[w] = {
                    "article": sku["article"],
                    "linear": lin,
                    "m2": sku["tyumen_m2"],
                }
            leftover = leftover or sku["leftover"]
        if not by_w:
            continue
        widths = sorted(by_w)
        stock = {str(w if w != int(w) else int(w) if abs(w - int(w)) < 1e-6 else w): by_w[w]["linear"] for w in widths}
        # normalize keys like 3.5 and 3
        stock_by_width = {}
        articles_by_width = {}
        for w in widths:
            key = str(w)
            if key.endswith(".0"):
                key = key[:-2]
            stock_by_width[key] = by_w[w]["linear"]
            articles_by_width[key] = by_w[w]["article"]
        width_nums = []
        for w in widths:
            width_nums.append(int(w) if abs(w - int(w)) < 1e-6 else w)

        price = 0
        for w in widths:
            a = by_w[w]["article"]
            if a in prices:
                price = prices[a]
                break
        slug = slugify(base)
        if slug in used_slugs:
            slug = f"{slug}-{articles_by_width[str(width_nums[0])].lower()}"
        used_slugs.add(slug)
        total_lin = sum(stock_by_width.values())
        promo = None
        if leftover or total_lin < 8:
            promo = "Остаток"
        title = base if base.lower().startswith("линолеум") else f"Линолеум {base}"
        title = re.sub(r"\s+", " ", title).strip()
        cat = category_of(base)
        tex = texture_of(base)
        rec: dict = {
            "kind": "linoleum",
            "id": f"ln-{slugify(articles_by_width[str(width_nums[0])])}",
            "slug": slug,
            "title": title,
            "brand": brand,
            "collection": str(collection),
            "design": str(design),
            "colorCode": design.split()[-1] if design else "",
            "category": cat,
            "utp": (
                f"На складе в Тюмени. Ширины {', '.join(f'{w} м' for w in width_nums)}. "
                "Нарезка в день заказа, доставка по городу бесплатно."
            ),
            "wearClass": 32 if cat.startswith("kommer") or cat == "polukommercheskiy" else 23,
            "thicknessMm": 2.5,
            "wearLayerMm": 0.4 if "supreme" in base.lower() else 0.25,
            "abrasionGroup": "T" if "supreme" in base.lower() else "P",
            "fireClass": "KM2" if cat.startswith("kommer") else "KM3",
            "impactNoiseDb": 17,
            "backing": "foam-pvc",
            "underfloorHeating": True,
            "widthsM": width_nums,
            "patternRepeatCm": 100,
            "pattern": "wood" if tex.startswith("oak") or tex == "walnut" else "abstract",
            "texture": tex,
            "pricePerM2": price,
            "stockByWidth": stock_by_width,
            "articlesByWidth": articles_by_width,
            "purposes": ["kitchen", "corridor", "living"],
            "specsUnknown": True,
            "priceOnRequest": price <= 0,
        }
        if promo:
            rec["promoLabel"] = promo
        products.append(rec)

    live = {
        "asOf": date.today().isoformat(),
        "designs": len(products),
        "skus": sum(len(p["widthsM"]) for p in products),
        "tyumenLinearM": round(sum(sum(p["stockByWidth"].values()) for p in products), 1),
        "items": [
            {
                "slug": p["slug"],
                "title": p["title"],
                "stockByWidth": p["stockByWidth"],
                "widthsM": p["widthsM"],
                "articlesByWidth": p.get("articlesByWidth") or {},
            }
            for p in products
        ],
    }
    OUT_LIVE.parent.mkdir(parents=True, exist_ok=True)
    OUT_LIVE.write_text(json.dumps(live, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if OUT_TS.parent.exists():
        lines = [
            'import type { LinoleumProduct } from "./types";',
            "",
            f'export const STOCK_AS_OF = "{date.today().isoformat()}";',
            f"export const STOCK_COUNT = {len(products)};",
            "",
            "export const LINOLEUM: LinoleumProduct[] = " + json.dumps(products, ensure_ascii=False, indent=2) + ";",
            "",
        ]
        ts = "\n".join(lines)
        ts = ts.replace(": null", ": undefined")
        OUT_TS.write_text(ts, encoding="utf-8")
    if OUT_JSON.parent.exists():
        OUT_JSON.write_text(
            json.dumps(
                {
                    "asOf": date.today().isoformat(),
                    "designs": len(products),
                    "skus": sum(len(p["widthsM"]) for p in products),
                    "tyumenLinearM": round(sum(sum(p["stockByWidth"].values()) for p in products), 1),
                    "tyumenCols": tyumen_cols,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
    print(f"wrote {len(products)} designs → {OUT_LIVE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
