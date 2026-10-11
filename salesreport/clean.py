"""Reading and cleaning order exports.

Every input file is mapped to the same set of columns, values are parsed
(Norwegian and English number and date formats), and rows that cannot be used
are set aside with a reason instead of being silently dropped.
"""

import csv
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

COLUMNS = ["order_id", "date", "city", "product", "quantity", "unit_price", "channel"]

# Header names seen in different exports, per canonical column.
# Matching ignores case, spaces and punctuation, so "Order ID" == "order_id".
COLUMN_ALIASES = {
    "order_id": ["order_id", "order number", "ordrenr", "ordrenummer"],
    "date": ["date", "order date", "dato", "ordredato"],
    "city": ["city", "customer_city", "by", "sted"],
    "product": ["product", "item", "produkt", "vare"],
    "quantity": ["quantity", "qty", "antall"],
    "unit_price": ["unit_price", "price", "price nok", "pris", "enhetspris"],
    "channel": ["channel", "sales channel", "kanal", "salgskanal"],
}

DATE_FORMATS = ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d.%m.%y"]


def _key(text):
    return re.sub(r"[^a-z0-9æøå]", "", str(text).lower())


_ALIAS_LOOKUP = {_key(alias): column for column, aliases in COLUMN_ALIASES.items() for alias in aliases}


def read_file(path):
    """Read one CSV or Excel export as text/objects, without guessing types."""
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        return pd.read_excel(path, dtype=object)

    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
    delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t").delimiter
    return pd.read_csv(path, sep=delimiter, dtype=str, encoding="utf-8-sig", keep_default_na=False)


def standardize(raw, source_name):
    """Rename columns to the canonical names and remember where each row came from."""
    mapping = {}
    for col in raw.columns:
        canonical = _ALIAS_LOOKUP.get(_key(col))
        if canonical and canonical not in mapping.values():
            mapping[col] = canonical

    missing = [c for c in COLUMNS if c not in mapping.values()]
    if missing:
        raise ValueError(
            f"{source_name}: could not find column(s) {', '.join(missing)}. "
            f"Columns in the file: {', '.join(map(str, raw.columns))}"
        )

    df = raw.rename(columns=mapping)[COLUMNS].copy()
    df["source_file"] = source_name
    df["source_row"] = range(2, len(df) + 2)  # row 1 is the header, as in Excel
    return df


def parse_number(value):
    """Parse '1 234,50', '1,234.50', '169.00', '429 kr' or a real number. None if not a number."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = re.sub(r"(?i)\b(kr|nok)\b|[\s  ]", "", str(value))
    if not text:
        return None
    if "," in text and "." in text:
        # The separator that comes last is the decimal one.
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        # Norwegian style: a single comma is the decimal separator.
        text = text.replace(",", ".") if text.count(",") == 1 else text.replace(",", "")
    elif re.fullmatch(r"\d{1,3}(\.\d{3}){2,}", text):
        text = text.replace(".", "")

    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value):
    """Parse a date from Excel or from text in common Norwegian/ISO formats. None if impossible."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def product_key(name):
    """Comparable form of a product name: 'ETHIOPIA  Yirgacheffe 250 g' -> 'ethiopia yirgacheffe 250g'."""
    text = " ".join(str(name).split()).lower()
    return re.sub(r"(\d)\s+(g|kg|ml|l)\b", r"\1\2", text)


def unify_products(names):
    """Map every spelling to the most common clean spelling of the same product."""
    cleaned = names.map(lambda n: " ".join(str(n).split()))
    keys = cleaned.map(product_key)
    display = {}
    for key, group in cleaned.groupby(keys):
        display[key] = group.value_counts().idxmax()
    return keys.map(display)


def clean(frames):
    """Combine standardized frames. Returns (clean rows, rows set aside with a reason)."""
    df = pd.concat(frames, ignore_index=True)

    for col in ("order_id", "city", "channel", "product"):
        df[col] = df[col].map(lambda v: " ".join(str(v).split()) if pd.notna(v) else "")

    df["date"] = df["date"].map(parse_date)
    df["quantity"] = pd.to_numeric(df["quantity"].map(parse_number), errors="coerce")
    df["unit_price"] = pd.to_numeric(df["unit_price"].map(parse_number), errors="coerce")

    reasons = pd.Series("", index=df.index, dtype=object)
    checks = [
        (df["order_id"] == "", "missing order number"),
        (df["product"] == "", "missing product"),
        (df["date"].isna(), "date missing or not a valid date"),
        (df["quantity"].isna() | (df["quantity"] <= 0), "quantity missing or not a positive number"),
        (df["unit_price"].isna() | (df["unit_price"] <= 0), "price missing or not a positive number"),
    ]
    for mask, reason in checks:
        reasons[mask & (reasons == "")] = reason

    valid = df[reasons == ""].copy()
    valid["product"] = unify_products(valid["product"])

    # The same order exported twice: identical order number, product, quantity and price.
    dup_cols = ["order_id", "product", "quantity", "unit_price"]
    duplicated = valid.duplicated(subset=dup_cols, keep="first")
    first_seen = valid.drop_duplicates(subset=dup_cols).set_index(dup_cols)[["source_file", "source_row"]]
    for idx in valid.index[duplicated]:
        row = valid.loc[idx]
        orig = first_seen.loc[tuple(row[c] for c in dup_cols)]
        reasons[idx] = f"duplicate of {orig['source_file']} row {orig['source_row']}"
    valid = valid[~duplicated].copy()

    valid["quantity"] = valid["quantity"].astype(int)
    valid["revenue"] = valid["quantity"] * valid["unit_price"]
    valid["month"] = valid["date"].map(lambda d: d.strftime("%Y-%m"))
    valid = valid.sort_values(["date", "order_id"]).reset_index(drop=True)

    issues = df.loc[reasons != "", ["source_file", "source_row", "order_id", "product"]].copy()
    issues.insert(2, "reason", reasons[reasons != ""])
    issues = issues.sort_values(["source_file", "source_row"]).reset_index(drop=True)
    return valid, issues


def load_folder(folder):
    """Read every CSV/Excel file in a folder. Returns (standardized frames, file names)."""
    files = sorted(p for p in Path(folder).iterdir() if p.suffix.lower() in (".csv", ".xlsx", ".xlsm"))
    if not files:
        raise FileNotFoundError(f"No .csv or .xlsx files found in {folder}")
    frames = [standardize(read_file(p), p.name) for p in files]
    return frames, [p.name for p in files]
