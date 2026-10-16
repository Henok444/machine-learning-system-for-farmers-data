import pandas as pd
import argparse
import os
import re
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path



BASE_DIR = Path(__file__).resolve().parents[1]
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
HEADER_SCAN_ROWS = 30
NEW_COLUMN = "__woreda_new__"



COLUMN_RULES = {
    "phone": (["phone", "contact", "mobile", "telephone"], ["emergency"]),
    "woreda": (["woreda", "wereda"], []),
    "first_name": (["firstname", "fname"], []),
    "middle_name": (["middlename"], []),
    "last_name": (["lastname", "lname", "surname"], []),
    "name": (["name", "farmer"], ["first", "middle", "last", "emergency"]),
    "gender": (["gender", "sex"], []),
    "region": (["region"], []),
    "zone": (["zone"], []),
    "kebele": (["kebele"], []),
    "status": (["status"], ["fayda"]),
    "value_chain": (["valuechain"], []),
    "land": (["land"], []),
    "seed": (["seed", "quantity"], []),
    "parcel": (["parcel"], []),
    "variety": (["variety", "vareity"], []),
    "remark": (["remark"], []),
}

SHEET_NAME_IS_WOREDA = True
FILL_BLANK_WOREDA_WHEN_SHEET_HAS_ONE = True


class UserError(Exception):
    pass


def key_of(name):
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def canonical_of(header):
    key = key_of(header)
    if not key or len(key) > 40:
        return None
    for canon, (words, banned) in COLUMN_RULES.items():
        if any(w in key for w in words) and not any(b in key for b in banned):
            return canon
    return None


def clean_text(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def normalize_phone(value):
    text = clean_text(value)
    if not text:
        return None
    if re.fullmatch(r"\d+(\.\d+)?[eE]\+?\d+", text):
        text = str(int(float(text)))
    text = re.sub(r"\.0+$", "", text)
    digits = re.sub(r"\D", "", text)
    if not digits:
        return None
    if digits.startswith("00251") and len(digits) == 14:
        digits = digits[5:]
    elif digits.startswith("251") and len(digits) == 12:
        digits = digits[3:]
    elif digits.startswith("0") and len(digits) == 10:
        digits = digits[1:]
    if len(digits) == 9 and digits[0] in "79":
        return "0" + digits
    return digits


def phone_keys(value):
    text = clean_text(value)
    if not text:
        return []
    keys = [normalize_phone(part) for part in re.split(r"[/,;&|\n]", text)]
    return [k for k in keys if k]


def sheet_key(name):
    return (clean_text(name) or "").lower()


def excel_row(index_label):
    return int(index_label) + 1


def short(error):
    message = (str(error).strip().splitlines() or [""])[0][:150]
    if message:
        return f"{type(error).__name__}: {message}"
    return type(error).__name__


@dataclass
class SheetInfo:
    path: Path
    name: str
    raw: pd.DataFrame
    header_row: int
    cols: set
    role: str


class Table:
    def __init__(self, df, display):
        self.df = df
        self.display = display


def read_sheets(path):
    try:
        return pd.read_excel(path, sheet_name=None, header=None, dtype=object, engine="openpyxl")
    except Exception as e:
        raise UserError(
            f"{path.name} cannot be read ({short(e)}). "
            "Is it a real .xlsx file, not password-protected, and not corrupted?"
        )


def find_header(df):
    best_row, best_cols = 0, set()
    for i in range(min(len(df), HEADER_SCAN_ROWS)):
        cols = {c for c in (canonical_of(v) for v in df.iloc[i] if clean_text(v)) if c}
        if len(cols) > len(best_cols):
            best_row, best_cols = i, cols
    return best_row, best_cols


def role_of(cols):
    if "phone" not in cols:
        return None
    if {"first_name", "last_name"} & cols:
        return "second"
    if "woreda" in cols or len(cols & {"name", "gender", "kebele", "region", "zone", "status"}) >= 2:
        return "first"
    return None


def build_table(info):
    raw_headers = [clean_text(v) for v in info.raw.iloc[info.header_row]]
    internal, display, seen = [], {}, set()
    for n, header in enumerate(raw_headers):
        unique = header or f"__blank_{n}"
        while unique in seen:
            unique += f"_{n}"
        seen.add(unique)
        internal.append(unique)
        display[unique] = header or ""
    body = info.raw.iloc[info.header_row + 1:].copy()
    body.columns = internal
    body = body.dropna(how="all")
    empty_blank = [c for c in internal if display[c] == "" and body[c].isna().all()]
    return Table(body.drop(columns=empty_blank), display)


def find_col(table, canonical):
    for col in table.df.columns:
        if canonical_of(table.display[col]) == canonical:
            return col
    return None


def file_role(sheets):
    first = sum(s.role == "first" for s in sheets)
    second = sum(s.role == "second" for s in sheets)
    if first > second:
        return "first"
    if second > first:
        return "second"
    return "mixed" if first else None


def safe_sheet_names(names):
    out, used = [], set()
    for n in names:
        base = re.sub(r"[\[\]\:\*\?/\\]", "_", str(n)).strip()[:31] or "Sheet"
        name, i = base, 2
        while name.lower() in used:
            suffix = f"_{i}"
            name, i = base[:31 - len(suffix)] + suffix, i + 1
        used.add(name.lower())
        out.append(name)
    return out


def write_output(sheets_out, final_path):
    tmp = final_path.with_name(f".~{final_path.name}")
    names = safe_sheet_names([s[0] for s in sheets_out])
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
            for name, (_, df, headers) in zip(names, sheets_out):
                df.to_excel(writer, sheet_name=name, index=False, header=headers)
                ws = writer.sheets[name]
                if headers is not False:
                    ws.freeze_panes = "A2"
                    for i, col in enumerate(df.columns, start=1):
                        sample = [len(str(v)) for v in df[col].head(200) if clean_text(v)]
                        width = max([len(headers[i - 1]), *sample]) + 2
                        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(width, 45)
        check = pd.read_excel(tmp, sheet_name=None, header=None, dtype=object, engine="openpyxl")
        for name, (_, df, headers) in zip(names, sheets_out):
            expected = len(df) + (0 if headers is False else 1)
            if name not in check or len(check[name]) != expected:
                raise UserError("The output file failed its read-back check, so it was discarded.")
        os.replace(tmp, final_path)
    finally:
        try:
            tmp.unlink()
        except OSError:
            pass


def free_output_path(second_path):
    path = OUTPUT_DIR / f"{second_path.stem}_with_woreda.xlsx"
    if path.exists():
        path = OUTPUT_DIR / f"{second_path.stem}_with_woreda_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    return path


def line(char="="):
    print(char * 70)


def run(check_only):
    line()
    print(f" Woreda adder   {datetime.now():%Y-%m-%d %H:%M:%S}" + ("   [CHECK ONLY]" if check_only else ""))
    line()
    # error handling 1 ================================
    try:
        INPUT_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        raise UserError(f"Cannot create the input/output folders next to the script ({short(e)}).")
    #=============================================================================================

    print(f"\n[1/4] Looking in {INPUT_DIR}")
    try:
        everything = sorted(p for p in INPUT_DIR.iterdir() if p.is_file() and not p.name.startswith((".", "~$")))
    except OSError as e:
        raise UserError(f"Cannot read the input folder ({short(e)}).")
    excel_files = [p for p in everything if p.suffix.lower() in EXCEL_SUFFIXES]
    for p in everything:
        if p not in excel_files:
            print(f"  ignored   : {p.name}  (only .xlsx files are supported)")
    if not excel_files:
        raise UserError(
            "The input folder has no .xlsx files.\n"
            f"  Put the first file and the second file in: {INPUT_DIR}"
        )
    print(f"  found     : {len(excel_files)} Excel file(s)")

    print("\n[2/4] Reading every sheet and recognising the two files")
    workbooks, problems = {}, []
    for path in excel_files:
        try:
            sheets = read_sheets(path)
        except UserError as e:
            problems.append(str(e))
            print(f"  {path.name}\n      UNREADABLE (corrupted, password-protected or not a real .xlsx)")
            continue
        infos = []
        for name, df in sheets.items():
            header_row, cols = find_header(df)
            infos.append(SheetInfo(path, str(name), df, header_row, cols, role_of(cols)))
        workbooks[path] = infos
        role = file_role(infos)
        label = {"first": "FIRST file", "second": "SECOND file", "mixed": "mixed (first + second sheets)"}.get(role, "not recognised")
        print(f"  {path.name}\n      -> {label}   ({len(infos)} sheet(s))")
        for s in infos:
            tag = {"first": "first ", "second": "second", None: "other "}[s.role]
            extra = ""
            if s.role is None:
                extra = f"  found: {', '.join(sorted(s.cols)) or 'no known columns'}"
            elif s.role == "first" and "woreda" not in s.cols:
                extra = "  (no Woreda column)"
            print(f"        [{tag}] '{s.name}'  header row {s.header_row + 1}{extra}")

    roles = {p: file_role(i) for p, i in workbooks.items()}
    firsts = [p for p, r in roles.items() if r == "first"]
    seconds = [p for p, r in roles.items() if r == "second"]
    mixed = [p for p, r in roles.items() if r == "mixed"]
    if not firsts and len(mixed) == 1 and len(seconds) == 1:
        firsts = mixed
    elif not seconds and len(mixed) == 1 and len(firsts) == 1:
        seconds = mixed
    if len(firsts) != 1 or len(seconds) != 1:
        reasons = list(problems)
        if not firsts:
            reasons.append("No FIRST file found (sheets with a phone column and a Woreda column).")
        if not seconds:
            reasons.append("No SECOND file found (sheets with phone and first_name/last_name columns).")
        if len(firsts) > 1:
            reasons.append("More than one FIRST file: " + ", ".join(p.name for p in firsts))
        if len(seconds) > 1:
            reasons.append("More than one SECOND file: " + ", ".join(p.name for p in seconds))
        reasons.append("Keep exactly the two files you want in the input folder. "
                       "If the files are right, a column name is probably not recognised: see COLUMN_RULES at the top of add_woreda.py.")
        raise UserError("\n  ".join(reasons))
    first_path, second_path = firsts[0], seconds[0]
    extra_files = [p for p in excel_files if p not in (first_path, second_path)]
    print(f"\n  FIRST  = {first_path.name}\n  SECOND = {second_path.name}")

    print("\n[3/4] Matching by phone number")
    by_sheet, everywhere, conflicts = {}, {}, set()
    blank_woreda_phones = set()
    usable_total = unusable_total = 0
    print("  FIRST file sheets used:")
    for info in workbooks[first_path]:
        if info.role != "first":
            print(f"      '{info.name}': skipped (not a farmers list)")
            continue
        table = build_table(info)
        phone_col, woreda_col = find_col(table, "phone"), find_col(table, "woreda")
        if woreda_col is None and not SHEET_NAME_IS_WOREDA:
            print(f"      '{info.name}': skipped (no Woreda column and SHEET_NAME_IS_WOREDA is False)")
            continue
        fallback = clean_text(info.name)
        local = by_sheet.setdefault(sheet_key(info.name), {})
        raw_woredas = [clean_text(v) for v in table.df[woreda_col]] if woreda_col is not None else None
        single = None
        if raw_woredas is not None and FILL_BLANK_WOREDA_WHEN_SHEET_HAS_ONE:
            distinct = {x.lower(): x for x in raw_woredas if x}
            if len(distinct) == 1:
                single = next(iter(distinct.values()))
        usable = inferred = 0
        for pos, phone_value in enumerate(table.df[phone_col]):
            keys = phone_keys(phone_value)
            guessed = False
            if raw_woredas is None:
                woreda = fallback
            else:
                woreda = raw_woredas[pos]
                if not woreda and single:
                    woreda, guessed = single, True
            if not keys:
                continue
            if not woreda:
                blank_woreda_phones.update(keys)
                continue
            usable += 1
            inferred += guessed
            for k in keys:
                if k in everywhere and everywhere[k].lower() != woreda.lower():
                    conflicts.add(k)
                everywhere.setdefault(k, woreda)
                local.setdefault(k, woreda)
        usable_total += usable
        unusable_total += len(table.df) - usable
        note = "" if woreda_col is not None else f"   (no Woreda column: sheet name '{fallback}' used as woreda)"
        if inferred:
            note = f"   ({inferred} blank Woreda cells filled with the sheet's only woreda '{single}')"
        print(f"      '{info.name}': {len(table.df)} rows, {usable} with phone + woreda{note}")
    if not everywhere:
        raise UserError("The first file has no row with both a phone number and a woreda.")
    print(f"  phone numbers available from the first file: {len(everywhere)}")

    print("  SECOND file sheets:")
    sheets_out, grand_total, grand_filled, grand_kept, unmatched_total = [], 0, 0, 0, 0
    for info in workbooks[second_path]:
        if info.role != "second":
            print(f"      '{info.name}': not a farmers table, copied unchanged")
            sheets_out.append((info.name, info.raw.dropna(how="all"), False))
            continue
        table = build_table(info)
        s_phone, existing_col = find_col(table, "phone"), find_col(table, "woreda")
        local = by_sheet.get(sheet_key(info.name), {})
        results, filled, kept, unmatched = [], 0, 0, []
        for idx, phone_value in zip(table.df.index, table.df[s_phone]):
            current = clean_text(table.df.at[idx, existing_col]) if existing_col is not None else None
            if current:
                results.append(current)
                kept += 1
                continue
            keys = phone_keys(phone_value)
            woreda = next((local[k] for k in keys if k in local), None) or next(
                (everywhere[k] for k in keys if k in everywhere), None)
            results.append(woreda)
            if woreda:
                filled += 1
            else:
                if not keys:
                    why = "no phone number"
                elif any(k in blank_woreda_phones for k in keys):
                    why = "phone is in first file but its woreda is empty there"
                else:
                    why = "phone not in first file"
                unmatched.append((excel_row(idx), clean_text(phone_value), why))

        series = pd.Series(results, index=table.df.index, dtype=object)
        if existing_col is not None:
            table.df[existing_col] = series
        else:
            position = len(table.df.columns)
            kebele_col = find_col(table, "kebele")
            if kebele_col is not None:
                position = list(table.df.columns).index(kebele_col) + 1
            table.df.insert(position, NEW_COLUMN, series)
            table.display[NEW_COLUMN] = "woreda"
        sheets_out.append((info.name, table.df, [table.display[c] for c in table.df.columns]))

        total = len(table.df)
        grand_total += total
        grand_filled += filled
        grand_kept += kept
        unmatched_total += len(unmatched)
        print(f"      '{info.name}': {total} rows -> {filled} filled"
              + (f", {kept} already had a woreda" if kept else "")
              + f", {len(unmatched)} no match")
        for row, phone, why in unmatched[:3]:
            print(f"          row {row}: {phone or '-'}  ({why})")
        if len(unmatched) > 3:
            print(f"          ... and {len(unmatched) - 3} more")

    if grand_total == 0:
        raise UserError(f"{second_path.name} has no data rows.")
    if grand_filled == 0 and grand_kept == 0:
        raise UserError(
            f"None of the {grand_total} phone numbers in the second file were found in the first file.\n"
            "  Check that these two files belong together. Nothing was written or removed."
        )
    print(f"  TOTAL: {grand_filled + grand_kept} of {grand_total} rows have a woreda, {unmatched_total} without")
    if unusable_total:
        print(f"  first-file rows that could not be used (no phone or no woreda): {unusable_total}")
    if conflicts:
        print(f"  WARNING: {len(conflicts)} phone number(s) have different woredas in the first file "
              "- the same-named sheet was preferred, otherwise the first one found")

    if check_only:
        line()
        print(" CHECK ONLY - nothing was written and no file was removed.")
        line()
        return 0

    print("\n[4/4] Writing the output")
    output_path = free_output_path(second_path)
    try:
        write_output(sheets_out, output_path)
    except PermissionError:
        raise UserError(f"Cannot write {output_path.name}: permission denied (is the file open in Excel?).")
    except OSError as e:
        raise UserError(f"Cannot write the output file ({short(e)}).")
    print(f"  saved     : {output_path}   ({len(sheets_out)} sheets)")

    removed, stuck = [], []
    for path in (first_path, second_path):
        try:
            path.unlink()
            removed.append(path.name)
        except OSError as e:
            stuck.append(f"{path.name} ({short(e)})")
    print(f"  removed from input : {', '.join(removed) if removed else 'nothing'}")
    if stuck:
        print("  WARNING: could not remove: " + "; ".join(stuck))
        print("           Close them in Excel and delete them by hand before the next run.")
    if extra_files:
        print("  left in input      : " + ", ".join(p.name for p in extra_files) + "  (not part of this run)")

    line()
    print(" DONE" + ("  (with warnings)" if (stuck or conflicts or unmatched_total) else "  - everything matched"))
    print(f" {grand_filled + grand_kept}/{grand_total} rows have a woreda   |   output: {output_path.name}")
    line()
    return 0


def main():
    parser = argparse.ArgumentParser(description="Add a woreda column to the second file.")
    parser.add_argument("--check", action="store_true",
                        help="only show what would happen; write nothing, remove nothing")
    args = parser.parse_args()
    try:
        return run(args.check)
    except UserError as e:
        print(f"\nCANNOT CONTINUE: {e}")
        print("No input file was removed. Fix the problem and run the script again.")
        return 1
    except KeyboardInterrupt:
        print("\nStopped by you. No input file was removed.")
        return 130
    except Exception as e:
        print(f"\nUNEXPECTED ERROR: {short(e)}")
        print("No input file was removed. Details for debugging:")
        traceback.print_exc()
        return 1


#====================================create a system ==================================

BASE = Path(__file__).resolve().parents[1]

path = BASE / "data" / "cleaned"
output_path = BASE / "data" / "woreda_added"

list_of_files = [file for file in path.glob("*.xlsx")]




if __name__ == "__main__":
    sys.exit(main())