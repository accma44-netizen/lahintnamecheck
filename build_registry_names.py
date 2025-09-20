# -*- coding: utf-8 -*-
import pandas as pd, re, json, os
from pathlib import Path

INPUT_XLSX = "all.xlsx"              # ضع الملف في نفس المجلد
OUT_JSON   = "registry_names.json"   # الناتج

def normalize_ar(s: str) -> str:
    if pd.isna(s): return ""
    s = str(s)
    s = re.sub(r"[\u0617-\u061A\u064B-\u0652\u0657-\u065F\u0670\u06D6-\u06ED]", "", s)  # remove diacritics
    s = re.sub("[إأٱآا]", "ا", s)
    s = s.replace("ة","ه").replace("ى","ي").replace("ئ","ي").replace("ؤ","و").replace("ـ","")
    s = re.sub(r"[^0-9\u0621-\u063A\u064F-\u064A\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def guess_name_cols(df: pd.DataFrame):
    out = []
    for c in df.columns:
        lc = str(c).lower()
        cc = str(c)
        if ("name" in lc) or ("trade" in lc) or ("company" in lc) or ("brand" in lc) or \
           ("mark" in lc) or ("اسم" in cc) or ("التجاري" in cc) or ("السجل" in cc) or ("الشركة" in cc) or ("الكيان" in cc):
            out.append(c)
    if not out:
        for c in df.columns:
            if df[c].dtype == "object":
                out.append(c); break
    return out

def main():
    if not Path(INPUT_XLSX).exists():
        raise SystemExit(f"لم يتم العثور على {INPUT_XLSX}")

    rows = []
    xls = pd.ExcelFile(INPUT_XLSX)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
        cols = guess_name_cols(df)
        for c in cols:
            for v in df[c].dropna().astype(str):
                nm = v.strip()
                if not nm or nm.lower() in ("nan","none","null"): 
                    continue
                rows.append({"name_original": nm, "name_normalized": normalize_ar(nm),
                             "source": "commercial_registry", "subsource": sheet})

    # تنظيف وتوحيد
    out = pd.DataFrame(rows)
    out = out[out["name_normalized"].str.len() >= 2]
    out = out.drop_duplicates(subset=["name_normalized"]).sort_values("name_normalized").reset_index(drop=True)
    out.to_json(OUT_JSON, orient="records", force_ascii=False)
    print(f"تم إنشاء {OUT_JSON} بعدد سجلات: {len(out)}")

if __name__ == "__main__":
    main()
