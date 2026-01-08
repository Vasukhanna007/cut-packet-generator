# app.py
# Streamlit web app: Cut Packet Generator with Excel SUMIFS formulas
# - Dropdown shows "Base Product (count)" sorted by most-active first
# - Base Product extractor handles -, – or — and strips sizes/options properly
# - Section B totals use SUMIFS over SectionA table (live formulas)

import re
from io import BytesIO
from datetime import datetime, date
from typing import List, Tuple
import pandas as pd
import numpy as np
import streamlit as st

# =========================
# ====== Core Settings ====
# =========================

ACCESSORY_KEYWORDS = ["potli","bag","pouch","dupatta","scarf","shawl","stole","belt","cap","mask"]
ALPHA_ORDER = ["XXXS","XXS","XS","S","M","L","XL","XXL","XXXL","FREE SIZE"]
SIZE_TOKEN = r"(XXXL|XXL|XL|XXS|XS|S|M|L|XXXS|FREE SIZE|[2-5]\d)"

# =========================
# ====== CSV Helpers ======
# =========================

def _read_orders_csvlike(file) -> pd.DataFrame:
    if isinstance(file, BytesIO):
        file.seek(0)
    for enc in ["utf-8-sig", "utf-8", "cp1252"]:
        try:
            return pd.read_csv(file, encoding=enc)
        except Exception:
            if hasattr(file, "seek"): file.seek(0)
            continue
    raise RuntimeError("Could not read CSV. Try re-exporting from Shopify.")

def _find_col(df, possible_names):
    # exact match first
    for name in df.columns:
        for cand in possible_names:
            if cand and name.lower().strip() == cand.lower().strip():
                return name
    # contains match next
    for name in df.columns:
        for cand in possible_names:
            if cand and cand.lower() in name.lower():
                return name
    return None

def _detect_columns(df):
    cols = {}
    cols["order"] = _find_col(df, ["Name","Order Number","Order name","Order","name"])
    cols["sku"] = _find_col(df, ["Lineitem sku","SKU","Sku","lineitem sku"])
    cols["title"] = _find_col(df, ["Lineitem name","Product Name","Title","Product title","Lineitem title"])
    cols["variant"] = _find_col(df, ["Lineitem variant","Variant Title","Variant","Lineitem variant title"])
    cols["qty"] = _find_col(df, ["Lineitem quantity","Quantity","Qty"])
    cols["date"] = _find_col(df, ["Created at","Created At","Order Date","Processed at","Paid at","created_at"])
    # Prioritize lineitem fulfillment status over order-level fulfillment status
    # Try exact match first for "Lineitem fulfillment status" (most common in Shopify exports)
    cols["fulfillment"] = None
    for col_name in df.columns:
        if col_name.strip().lower() == "lineitem fulfillment status":
            cols["fulfillment"] = col_name
            break
    # Fallback to other variations
    if not cols["fulfillment"]:
        cols["fulfillment"] = _find_col(df, ["Lineitem fulfillment status", "Lineitem Fulfillment Status", 
                                             "lineitem fulfillment status", "Fulfillment Status", 
                                             "Fulfillment status", "fulfillment_status"])
    cols["notes"] = _find_col(df, ["Notes","Order Notes","note"])
    cols["prop_cols"] = [c for c in df.columns if "lineitem properties" in c.lower()]
    return cols

# =========================
# ====== Date Helpers =====
# =========================

def _parse_any_date(x):
    if x is None or (isinstance(x, float) and np.isnan(x)): return None
    s = str(x).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S %z","%Y-%m-%d %H:%M:%S",
                "%d-%m-%Y %H:%M","%d/%m/%Y %H:%M","%m/%d/%Y %H:%M",
                "%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, utc=False, errors="coerce")
        if pd.isna(dt): return None
        return pd.to_datetime(dt).to_pydatetime().replace(tzinfo=None)
    except Exception:
        return None

def _build_order_date_map(df, order_col: str):
    candidates = []
    for nm in ["Created at","Created At","Processed at","Paid at","Order Date",
               "created_at","processed_at","paid_at"]:
        if nm in df.columns:
            candidates.append(nm)
    if not candidates:
        s = pd.Series(dtype="datetime64[ns]")
        s.index.name = "Order#"; return s
    best = None
    for i, col in enumerate(candidates):
        parsed = pd.to_datetime(df[col], errors="coerce", utc=True)
        best = parsed if i == 0 else best.fillna(parsed)
    ist = best.dt.tz_convert("Asia/Kolkata").dt.tz_localize(None)
    s = pd.Series(ist.values, index=df[order_col].astype(str), name="_OrderDateLocal")
    s.index.name = "Order#"
    return s.groupby(level=0).first()

# =========================
# ====== Notes/Status =====
# =========================

def _build_order_notes_map(df, order_col: str, notes_col: str | None):
    if not notes_col or notes_col not in df.columns:
        s = pd.Series(dtype="object"); s.index.name = "Order#"; return s
    s = pd.Series(df[notes_col].astype(str).replace({"nan": ""}).values,
                  index=df[order_col].astype(str), name="Notes")
    s.index.name = "Order#"
    s = s.replace("", pd.NA).groupby(level=0).first().fillna("")
    return s

def _is_unfulfilled(val):
    """Return True if item is unfulfilled, False if fulfilled."""
    if pd.isna(val) or val is None: return True
    s = str(val).strip().lower()
    # If explicitly "fulfilled", it's NOT unfulfilled (should be filtered out)
    if s == "fulfilled":
        return False
    # Check for partial fulfillment - if it contains "fulfilled" but not "unfulfilled", it's fulfilled
    if "fulfilled" in s and "unfulfilled" not in s:
        return False
    # Otherwise, check for unfulfilled indicators
    return (s == "") or ("unfulfilled" in s) or ("pending" in s) or ("not fulfilled" in s)

def _is_unfulfilled_series(series):
    return series.apply(_is_unfulfilled)

# =========================
# === Size/Accessory NLP ==
# =========================

TOP_PATTERNS = [
    rf"\btop\b\s*[:\-]?\s*{SIZE_TOKEN}\b",
    rf"\b{SIZE_TOKEN}\b\s*[-/]*\s*\btop\b",
    rf"\bupper\b\s*[:\-]?\s*{SIZE_TOKEN}\b",
]
BOTTOM_PATTERNS = [
    rf"\b(bottom|btm|pant|pants|trouser|lower)\b\s*[:\-]?\s*{SIZE_TOKEN}\b",
    rf"\b{SIZE_TOKEN}\b\s*[-/]*\s*\b(bottom|btm|pant|pants|trouser|lower)\b",
]

def _combine_properties_text(row, prop_cols):
    parts = []
    for c in prop_cols:
        v = row.get(c, None)
        if pd.notna(v) and str(v).strip() != "":
            parts.append(str(v))
    return " | ".join(parts).strip().lower()

def _find_size_from_match(match_obj):
    if not match_obj: return None
    for g in match_obj.groups():
        if not g: continue
        g2 = str(g).upper().strip()
        if re.fullmatch(r"(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)", g2):
            return g2
    return None

def _find_top_size(text):
    for pat in TOP_PATTERNS:
        m = re.search(pat, text, re.I)
        s = _find_size_from_match(m)
        if s: return s
    return None

def _find_bottom_size(text):
    for pat in BOTTOM_PATTERNS:
        m = re.search(pat, text, re.I)
        s = _find_size_from_match(m)
        if s: return s
    return None

def _find_accessories(text):
    found = []
    for word in ACCESSORY_KEYWORDS:
        if re.search(rf"\b{re.escape(word)}\b", text, re.I):
            if re.search(rf"\b(no|without)\s+{re.escape(word)}\b", text, re.I):
                continue
            found.append(word.capitalize())
    # dedupe
    seen = set(); out=[]
    for x in found:
        if x not in seen:
            seen.add(x); out.append(x)
    return out

def _extract_from_variant_or_title(variant, title):
    """Extract a single size from variant or title, prioritizing size after ' - '."""
    text = f"{variant or ''} {title or ''}"
    
    # First, try to extract size after " - " (most common pattern like "Product - M")
    dash_size_match = re.search(r"-\s*(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)\b", text, re.I)
    if dash_size_match:
        size = dash_size_match.group(1).upper()
        # Validate it's a real size
        if size in ALPHA_ORDER or re.fullmatch(r"[2-5]\d", size):
            return size
    
    # Fallback: search through ALPHA_ORDER (but prefer longer matches first to avoid "S" matching in "XS")
    # Sort by length descending to match "XXXS" before "XS" before "S"
    sorted_sizes = sorted(ALPHA_ORDER, key=len, reverse=True)
    for s in sorted_sizes:
        if re.search(rf"\b{re.escape(s)}\b", text, re.I):
            return s
    
    # Try numeric sizes
    m = re.search(r"\b([2-5]\d)\b", text)
    if m: return m.group(1)
    return None

def _is_single_item_product(title, variant):
    """Detect if product is a single item (top-only, bottom-only, accessory) vs a set.
    
    Logic:
    1. If sizing format is "X / Y" (two sizes) → it's a set
    2. If sizing is single "X" → check keywords to determine top or bottom
    3. If no size found → it's an accessory
    """
    text = f"{title or ''} {variant or ''}"
    
    # First, check if it has two sizes in format "X / Y" - this indicates a set
    two_size_pattern = re.search(
        r"-\s*(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)\s*/\s*"
        r"(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)\b", text, re.I
    )
    if two_size_pattern:
        return None  # It's a set (has two sizes)
    
    # Check if it has a single size (after dash or in text)
    single_size_pattern = re.search(
        r"-\s*(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)\b", text, re.I
    )
    has_single_size = single_size_pattern is not None
    
    # If no size found, it's an accessory
    if not has_single_size:
        # Check for accessory keywords to confirm
        text_lower = text.lower()
        for acc in ACCESSORY_KEYWORDS:
            if re.search(rf"\b{re.escape(acc)}\b", text_lower, re.I):
                if not re.search(rf"\b(no|without)\s+{re.escape(acc)}\b", text_lower, re.I):
                    return "accessory"
        # If no accessory keyword but no size, still treat as accessory
        return "accessory"
    
    # Has single size - check keywords to determine if top or bottom
    text_lower = text.lower()
    
    # Keywords for top-only items
    top_only_keywords = ["top", "kurta", "shirt", "blouse", "tunic", "kurti", "cape", 
                         "jacket", "blazer", "coat"]
    # Keywords for bottom-only items
    bottom_only_keywords = ["bottom", "pant", "pants", "trouser", "trousers", "leggings", 
                           "palazzo", "salwar", "churidar", "dhoti", "farshi"]
    
    # Check for bottom keywords first (more specific)
    for kw in bottom_only_keywords:
        if re.search(rf"\b{re.escape(kw)}\b", text_lower, re.I):
            return "bottom"
    
    # Check for top keywords
    for kw in top_only_keywords:
        if re.search(rf"\b{re.escape(kw)}\b", text_lower, re.I):
            return "top"
    
    # Has size but no clear keyword - default to top (most common)
    return "top"

def _extract_two_sizes_from_variant_or_title(variant, title):
    """Return (top_size, bottom_size) using ' - X / Y' if present, else first two tokens."""
    def _tok2(txt):
        if not txt: return (None, None)
        T = str(txt).upper()
        # First, try to match " - X / Y" pattern (two sizes)
        m = re.search(
            r"-\s*(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)\s*/\s*"
            r"(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)\b", T, flags=re.I
        )
        if m: return m.group(1).upper(), m.group(2).upper()
        
        # Second, try to match " - X" pattern (single size after dash, most common)
        dash_single = re.search(r"-\s*(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)\b", T, flags=re.I)
        if dash_single:
            return dash_single.group(1).upper(), None
        
        # Last resort: find all sizes in the text (but prioritize those after dash)
        toks = re.findall(r"(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)", T, flags=re.I)
        toks = [t.upper() for t in toks]
        if len(toks) >= 2: return toks[0], toks[1]
        if len(toks) == 1: return toks[0], None  # Don't duplicate single size
        return (None, None)
    t1, t2 = _tok2(title)
    if t1 or t2: return t1, t2
    return _tok2(variant)

# =========================
# === Base Product name ===
# =========================

# tokens that indicate the right side is a size/options suffix
SUFFIX_SIZE = r"(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|FREE SIZE|[2-5]\d)"
SUFFIX_WORDS = r"(with|without|no|potli|bag|pouch|dupatta|scarf|shawl|belt|cap|mask|set of \d+)"
# split on ASCII hyphen, en dash, or em dash (one occurrence)
DASH_SPLIT_RE = re.compile(r"\s[-–—]\s", flags=re.I)

def _looks_like_suffix(s: str) -> bool:
    if not s: return False
    parts = re.split(r"[\/,\|\+\·\•]|(?:\s{2,})", s)
    parts = [p.strip().lower() for p in parts if p and p.strip()]
    if not parts:
        return False
    ok = 0
    for p in parts:
        if re.fullmatch(SUFFIX_SIZE, p.upper()):
            ok += 1; continue
        if re.search(rf"\b{SUFFIX_WORDS}\b", p, flags=re.I):
            ok += 1; continue
        if p in {"and", "-", "/", "&"}:
            ok += 1; continue
        if p in {"with", "without", "no"}:
            ok += 1; continue
    return ok >= max(1, len(parts) // 2)

def extract_base_product(title: str) -> str:
    if not isinstance(title, str):
        return str(title)
    t = title.strip()
    m = DASH_SPLIT_RE.split(t, maxsplit=1)
    if len(m) == 2:
        left, right = m[0].strip(), m[1].strip()
        if _looks_like_suffix(right):
            return left
    if " - " in t:
        left, right = t.split(" - ", 1)
        if _looks_like_suffix(right.strip()):
            return left.strip()
    return t

# =========================
# ===== Normalization =====
# =========================

def _normalize_components(df, cols) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        order = r.get(cols["order"], None) if cols["order"] else None
        sku = r.get(cols["sku"], None) if cols["sku"] else None
        title = r.get(cols["title"], None) if cols["title"] else None
        variant = r.get(cols["variant"], None) if cols["variant"] else None
        qty = r.get(cols["qty"], 1) if cols["qty"] else 1
        created = r.get(cols["date"], None) if cols["date"] else None
        fulfill = r.get(cols["fulfillment"], None) if cols["fulfillment"] else None

        try:
            qty = int(qty)
        except Exception:
            try: qty = int(float(qty))
            except Exception: qty = 1

        created_dt = _parse_any_date(created)
        prop_text = _combine_properties_text(r, cols["prop_cols"]) if cols["prop_cols"] else ""

        # Detect if this is a single item (top-only, bottom-only, accessory) vs a set
        single_item_type = _is_single_item_product(title, variant)
        
        # If it's an accessory, handle it separately (no sizes)
        if single_item_type == "accessory":
            # Find which accessory it is
            text = f"{title or ''} {variant or ''}".lower()
            found_accessory = None
            for acc in ACCESSORY_KEYWORDS:
                if re.search(rf"\b{re.escape(acc)}\b", text, re.I):
                    if not re.search(rf"\b(no|without)\s+{re.escape(acc)}\b", text, re.I):
                        found_accessory = acc.capitalize()
                        break
            
            # Also check properties for accessories
            prop_accessories = _find_accessories(prop_text)
            
            entries = []
            if found_accessory:
                entries.append((f"Accessory: {found_accessory}", None, qty))
            for acc in prop_accessories:
                entries.append((f"Accessory: {acc}", None, 1))
            
            # If no accessory found but detected as accessory, create generic entry
            if not entries and single_item_type == "accessory":
                # Try to extract accessory name from title
                title_str = str(title).lower() if pd.notna(title) else ""
                for acc in ACCESSORY_KEYWORDS:
                    if acc in title_str:
                        entries.append((f"Accessory: {acc.capitalize()}", None, qty))
                        break
            
            # Skip size detection for accessories - they have no sizes
            for comp, size, q in entries:
                rows.append({
                    "Date": created_dt,
                    "Order#": order,
                    "Product": str(title).strip() if pd.notna(title) else None,
                    "Component": comp,
                    "Size": None,  # Accessories have no size
                    "Qty": q,
                    "SKU": sku,
                    "Notes": None,  # filled later
                    "_FulfillmentStatus": fulfill,
                })
            continue  # Skip to next row
        
        # For non-accessories, proceed with size detection
        # Try sizes from variant/title, else from properties
        t_guess, b_guess = _extract_two_sizes_from_variant_or_title(variant, title)
        top_prop = _find_top_size(prop_text)
        bottom_prop = _find_bottom_size(prop_text)

        if t_guess or b_guess:
            top_size = t_guess
            bottom_size = b_guess
        else:
            top_size = top_prop
            bottom_size = bottom_prop

        # If single item detected, only assign size to the appropriate component
        if single_item_type == "top":
            # Top-only item: only create Top entry, ignore bottom
            if not top_size:
                # Extract single size if available
                single_size = _extract_from_variant_or_title(variant, title)
                if single_size:
                    top_size = single_size
            bottom_size = None  # Clear bottom size for top-only items
        elif single_item_type == "bottom":
            # Bottom-only item: only create Bottom entry, ignore top
            if not bottom_size:
                # Extract single size if available
                single_size = _extract_from_variant_or_title(variant, title)
                if single_size:
                    bottom_size = single_size
            top_size = None  # Clear top size for bottom-only items
        else:
            # Set or unknown: use both sizes if available
            # If looks like a Set and still missing sizes, fallback to a single token
            title_str = str(title) if pd.notna(title) else ""
            if (not top_size and not bottom_size) and ("set" in title_str.lower()):
                vt_size = _extract_from_variant_or_title(variant, title)
                if vt_size:
                    top_size = vt_size
                    bottom_size = vt_size

        accessories = _find_accessories(prop_text)

        entries = []
        if top_size: entries.append(("Top", top_size, qty))
        if bottom_size: entries.append(("Bottom", bottom_size, qty))
        for acc in accessories: entries.append((f"Accessory: {acc}", None, 1))

        for comp, size, q in entries:
            rows.append({
                "Date": created_dt,
                "Order#": order,
                "Product": str(title).strip() if pd.notna(title) else None,
                "Component": comp,
                "Size": size,
                "Qty": q,
                "SKU": sku,
                "Notes": None,  # filled later
                "_FulfillmentStatus": fulfill,
            })

    expected = ["Date","Order#","Product","Component","Size","Qty","SKU","Notes","_FulfillmentStatus"]
    if not rows:
        return pd.DataFrame({c: [] for c in expected})
    out = pd.DataFrame(rows)
    for c in expected:
        if c not in out.columns:
            out[c] = pd.NaT if c == "Date" else pd.Series(dtype="object")
    return out

# =========================
# ===== Generator Core ====
# =========================

def generate_cut_packet_generic_df(
    df_in: pd.DataFrame,
    base_products: List[str],
    start_date: str | date | None = None,
    end_date: str | date | None = None,
    only_unfulfilled: bool = True,
    exclude_cancel: bool = True,
    last_3_months_default: bool = True,
    filter_sizes_in_sectionA: List[str] | None = None,
    size_cols_override_for_sectionB: List[str] | None = None,
    min_age_days: int | None = None
) -> Tuple[pd.DataFrame, List[str], List[str], list]:

    df = df_in.copy()
    df.columns = [c.strip() for c in df.columns]
    cols = _detect_columns(df)

    order_col = cols.get("order")
    order_date_map = _build_order_date_map(df, order_col) if order_col else None
    order_notes_map = _build_order_notes_map(df, order_col, cols.get("notes")) if order_col else None

    # Mark cancelled orders to exclude
    cancelled_ids = set()
    if exclude_cancel and cols.get("notes") and cols.get("order") and cols["notes"] in df.columns:
        cancelled_ids = set(
            df[df[cols["notes"]].astype(str).str.contains("cancel", case=False, na=False)][cols["order"]]
            .astype(str).tolist()
        )

    norm = _normalize_components(df, cols)

    # Attach order-level Notes (first note per order)
    if order_notes_map is not None and not order_notes_map.empty and "Order#" in norm.columns:
        norm["Notes"] = norm["Order#"].astype(str).map(order_notes_map).fillna("")

    # Remove cancelled
    if cancelled_ids and "Order#" in norm.columns:
        norm = norm[~norm["Order#"].astype(str).isin(cancelled_ids)].copy()

    # Local IST date
    raw_dates = pd.to_datetime(norm.get("Date"), errors="coerce", utc=True)
    norm["_DateLocal"] = raw_dates.dt.tz_convert("Asia/Kolkata").dt.tz_localize(None)

    # Fill missing dates from order-level map
    if order_date_map is not None and not order_date_map.empty and "Order#" in norm.columns:
        idx = norm["Order#"].astype(str)
        fill_vals = order_date_map.reindex(idx).values
        missing = norm["_DateLocal"].isna()
        norm.loc[missing, "_DateLocal"] = fill_vals[missing]

    # Only unfulfilled - filter out fulfilled items
    if only_unfulfilled and "_FulfillmentStatus" in norm.columns:
        # Keep only items where _is_unfulfilled returns True (i.e., exclude fulfilled items)
        unfulfilled_mask = _is_unfulfilled_series(norm["_FulfillmentStatus"])
        norm = norm[unfulfilled_mask].copy()

    # Last 3 months by default
    if last_3_months_default:
        now_ist_naive = pd.Timestamp.now(tz="Asia/Kolkata").tz_localize(None)
        cutoff_default = now_ist_naive - pd.Timedelta(days=90)
        date_local = norm["_DateLocal"]
        keep_mask = date_local.isna() | (date_local >= cutoff_default)
        norm = norm[keep_mask].copy()

    # Date overrides
    if start_date:
        sd = pd.Timestamp(start_date)
        norm = norm[norm["_DateLocal"].isna() | (norm["_DateLocal"] >= sd)]
    if end_date:
        ed = pd.Timestamp(end_date)
        norm = norm[norm["_DateLocal"].isna() | (norm["_DateLocal"] <= ed)]
    
    # Minimum age filter (show only orders older than X days)
    if min_age_days is not None and min_age_days > 0:
        now_ist_naive = pd.Timestamp.now(tz="Asia/Kolkata").tz_localize(None)
        min_age_cutoff = now_ist_naive - pd.Timedelta(days=min_age_days)
        date_local = norm["_DateLocal"]
        # Keep orders that are older than min_age_days (date < cutoff)
        keep_mask = date_local.isna() | (date_local < min_age_cutoff)
        norm = norm[keep_mask].copy()

    # BaseProduct
    norm["BaseProduct"] = norm["Product"].astype(str).map(extract_base_product)

    # Filter by selected base products (skip if empty and using age filter)
    base_set = set([bp.strip() for bp in base_products if bp and bp.strip() != ""])
    if len(base_set) > 0:
        sub = norm[norm["BaseProduct"].isin(base_set)].copy()
    else:
        # If no base products selected and age filter is active, include all products
        if min_age_days is not None and min_age_days > 0:
            sub = norm.copy()
        else:
            # No base products and no age filter - return empty
            sub = norm.iloc[0:0].copy()

    matched_titles = sorted(sub["Product"].dropna().astype(str).unique().tolist())

    cols_out = ["Date","Order#","Product","Component","Size","Qty","SKU","Notes"]
    if sub.empty:
        return pd.DataFrame(columns=cols_out), [], [a.capitalize() for a in ACCESSORY_KEYWORDS], matched_titles

    # Section A (with display date)
    secA = sub[cols_out + ["_DateLocal"]].copy()
    secA["Date"] = secA["_DateLocal"].dt.strftime("%Y-%m-%d").fillna("")
    secA = secA.drop(columns=["_DateLocal"])

    # Optional size filter inside Section A
    if filter_sizes_in_sectionA and len(filter_sizes_in_sectionA) > 0:
        secA = secA[secA["Size"].astype(str).isin([str(s) for s in filter_sizes_in_sectionA])].copy()

    # Section B size columns
    if size_cols_override_for_sectionB and len(size_cols_override_for_sectionB) > 0:
        size_cols = [str(s) for s in size_cols_override_for_sectionB]
    else:
        sizes_present = secA["Size"].dropna().astype(str).tolist()
        alpha = [s for s in sizes_present if not re.fullmatch(r"[2-5]\d", str(s))]
        numeric = [int(s) for s in sizes_present if re.fullmatch(r"[2-5]\d", str(s))]
        alpha_sorted = [s for s in ALPHA_ORDER if s in alpha] + [s for s in alpha if s not in ALPHA_ORDER]
        numeric_sorted = [str(n) for n in sorted(set(numeric))]
        size_cols = alpha_sorted + numeric_sorted

    accessories = [a.capitalize() for a in ACCESSORY_KEYWORDS]
    return secA, size_cols, accessories, matched_titles

# =========================
# ===== Excel Writing =====
# =========================

def _xl_rc(r0, c0):
    """Zero-based -> absolute A1 like $A$1."""
    col = ""
    c = c0
    while True:
        col = chr(c % 26 + 65) + col
        c = c // 26 - 1
        if c < 0: break
    return f"${col}${r0+1}"

def write_excel_with_formulas_to_buffer(secA: pd.DataFrame, size_cols: List[str], accessories: List[str],
                                        product_label: str, start_date, end_date) -> BytesIO:
    out = BytesIO()

    # unique OrderNotes from Section A
    if not secA.empty and "Order#" in secA.columns and "Notes" in secA.columns:
        order_notes_df = secA.groupby("Order#", as_index=False)["Notes"].first()
    else:
        order_notes_df = pd.DataFrame({"Order#": [], "Notes": []})

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        # Section A as Table
        secA.to_excel(writer, index=False, sheet_name="SectionA_Orderwise")
        wb  = writer.book
        wsA = writer.sheets["SectionA_Orderwise"]

        widths = {"A":12,"B":16,"C":44,"D":12,"E":10,"F":8,"G":18,"H":36}
        for col_letter, w in widths.items():
            wsA.set_column(f"{col_letter}:{col_letter}", w)

        nrows, ncols = secA.shape
        table_range = f"A1:{chr(ord('A')+ncols-1)}{nrows+1}"
        wsA.add_table(table_range, {
            "name": "SectionA",
            "header_row": True,
            "style": "Table Style Light 9",
            "columns": [{"header": h} for h in secA.columns.tolist()]
        })

        # Section B with SUMIFS formulas
        wsB = wb.add_worksheet("SectionB_Totals")
        header_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        row_lbl_fmt = wb.add_format({"bold": True})
        cell_int_fmt = wb.add_format({"num_format": "0"})

        # headers
        wsB.write(0, 0, "")
        for j, s in enumerate(size_cols, start=1):
            wsB.write(0, j, s, header_fmt)

        # TOPS
        wsB.write(1, 0, "TOPS", row_lbl_fmt)
        for j in range(1, len(size_cols)+1):
            size_hdr = _xl_rc(0, j)
            f = f'=SUMIFS(SectionA[Qty],SectionA[Component],"Top",SectionA[Size],{size_hdr})'
            wsB.write_formula(1, j, f, cell_int_fmt)

        # BOTTOMS
        wsB.write(2, 0, "BOTTOMS", row_lbl_fmt)
        for j in range(1, len(size_cols)+1):
            size_hdr = _xl_rc(0, j)
            f = f'=SUMIFS(SectionA[Qty],SectionA[Component],"Bottom",SectionA[Size],{size_hdr})'
            wsB.write_formula(2, j, f, cell_int_fmt)

        # Accessories block
        start_row = 4
        if accessories:
            wsB.write(start_row-1, 0, "ACCESSORIES (count by type)", header_fmt)
        for i, acc in enumerate(accessories):
            r = start_row + i
            wsB.write(r, 0, acc, row_lbl_fmt)
            f = f'=SUMIFS(SectionA[Qty],SectionA[Component],"Accessory: {acc}")'
            wsB.write_formula(r, 1, f, cell_int_fmt)

        # OrderNotes
        if not order_notes_df.empty:
            order_notes_df.to_excel(writer, index=False, sheet_name="OrderNotes")
            wsN = writer.sheets["OrderNotes"]
            wsN.set_column("A:A", 18)
            wsN.set_column("B:B", 60)

        # README
        pd.DataFrame({
            "Field":["Base Product(s)","Filters Applied","How Section B works"],
            "Value":[
                product_label,
                f"only unfulfilled; <=3 months old; notes: 'cancel' excluded; "
                f"start_date={start_date or 'None'}, end_date={end_date or 'None'}",
                "Section B uses SUMIFS on the SectionA table. If you edit Size/Qty in SectionA, totals update automatically."
            ]
        }).to_excel(writer, index=False, sheet_name="README")

    out.seek(0)
    return out

# =========================
# ========= UI ============
# =========================

st.set_page_config(
    page_title="Cut Packet Generator", 
    page_icon="✂️", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("✂️ Cut Packet Generator (Streamlit)")
st.caption("Upload Shopify CSV → select Base Product(s) (dropdown shows count, most active first) → optional filters → download Excel. Section B uses SUMIFS so totals auto-update.")

with st.sidebar:
    st.header("Filters")
    only_unfulfilled = st.checkbox("Only unfulfilled", value=True)
    exclude_cancel = st.checkbox("Exclude orders with 'cancel' in Notes", value=True)
    last_3m = st.checkbox("Limit to last 3 months (default)", value=True)
    
    st.markdown("---")
    st.markdown("**Age Filter**")
    use_min_age = st.checkbox("Show only orders older than X days", value=False)
    min_age_days = None
    if use_min_age:
        min_age_days = st.number_input("Minimum age (days)", min_value=1, max_value=365, value=15, step=1)

uploaded = st.file_uploader("Upload Shopify Orders CSV", type=["csv"])

# Date range
c1, c2 = st.columns(2)
with c1:
    start = st.date_input("Start date (optional)", value=None)
with c2:
    end = st.date_input("End date (optional)", value=None)

# Advanced size controls
with st.expander("Advanced size controls (optional)", expanded=False):
    st.markdown("**Section A size filter** (leave blank to include all sizes):")
    size_alpha_choices = st.multiselect("Alpha sizes", ALPHA_ORDER, default=[])
    default_numeric = [str(n) for n in range(24, 47, 2)]
    size_numeric_choices = st.multiselect("Numeric sizes", default_numeric, default=[])
    secA_size_filter = size_alpha_choices + size_numeric_choices

    st.markdown("---")
    st.markdown("**Section B columns** (pick a standard set even if totals are 0):")
    use_standard_cols = st.checkbox("Use standard size columns below (instead of auto-detect)", value=False)
    if use_standard_cols:
        sb_alpha = st.multiselect("Section B alpha columns", ALPHA_ORDER, default=ALPHA_ORDER)
        sb_numeric = st.multiselect("Section B numeric columns", default_numeric, default=default_numeric)
        sectionB_size_cols_override = sb_alpha + sb_numeric
    else:
        sectionB_size_cols_override = None

base_choices = []
picked_bases = []
df0 = None
title_col0 = None

if uploaded:
    try:
        df0 = _read_orders_csvlike(uploaded)
        cols0 = _detect_columns(df0)
        title_col0 = cols0.get("title")
        if not title_col0:
            st.error("Could not find a product title/lineitem name column in your CSV.")
        else:
            # Build counts per BaseProduct from raw titles
            titles = df0[title_col0].dropna().astype(str).str.strip()
            base_series = titles.map(extract_base_product)
            counts = base_series.value_counts().to_dict()

            # Sort by count desc, then name asc
            bases_sorted = sorted(counts.keys(), key=lambda b: (-counts[b], b.lower()))
            label_for_base = {b: f"{b} ({counts.get(b, 0)})" for b in bases_sorted}
            base_for_label = {v: k for k, v in label_for_base.items()}

            # Show hint if age filter is enabled
            help_text = None
            if use_min_age and min_age_days:
                help_text = "Optional: Leave empty to include all products when age filter is active"
            
            picked_labels = st.multiselect(
                "Select Base Product(s) — most active shown first",
                options=[label_for_base[b] for b in bases_sorted],
                help=help_text
            )
            picked_bases = [base_for_label[lbl] for lbl in picked_labels]
    except Exception as e:
        st.error(f"Failed to read CSV: {e}")

# Enable button if base products selected OR if age filter is enabled
button_disabled = len(picked_bases) == 0 and not (use_min_age and min_age_days)

if uploaded and st.button("Generate Excel", type="primary", disabled=button_disabled):
    with st.spinner("Processing…"):
        try:
            secA, size_cols, accessories, matched_titles = generate_cut_packet_generic_df(
                df_in=df0,
                base_products=picked_bases,
                start_date=start if start else None,
                end_date=end if end else None,
                only_unfulfilled=only_unfulfilled,
                exclude_cancel=exclude_cancel,
                last_3_months_default=last_3m,
                filter_sizes_in_sectionA=(secA_size_filter if len(secA_size_filter)>0 else None),
                size_cols_override_for_sectionB=sectionB_size_cols_override,
                min_age_days=min_age_days
            )

            if len(matched_titles) == 0:
                if len(picked_bases) == 0:
                    st.warning("No orders matched the age filter criteria.")
                else:
                    st.warning("No lines matched those Base Product(s) with current filters.")
            else:
                with st.expander(f"Matched full titles ({len(matched_titles)})", expanded=False):
                    st.write(matched_titles)

            if secA.empty:
                st.warning("No matching unfulfilled orders (respecting filters/date range).")
            else:
                st.subheader("Preview — Section A (first 50 rows)")
                st.dataframe(secA.head(50), use_container_width=True)

                if len(picked_bases) > 0:
                    product_label = ", ".join(picked_bases[:5]) + (" …" if len(picked_bases) > 5 else "")
                else:
                    product_label = "All Products (age filter only)"
                xls = write_excel_with_formulas_to_buffer(
                    secA=secA,
                    size_cols=size_cols if size_cols else ALPHA_ORDER,   # fallback so sheet isn't empty
                    accessories=accessories,
                    product_label=product_label,
                    start_date=start,
                    end_date=end
                )
                st.download_button(
                    label="⬇️ Download Excel",
                    data=xls,
                    file_name="cut_packet_OUTPUT.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error: {e}")
else:
    st.info("Upload your Shopify CSV, then pick one or more Base Product(s) from the dropdown (counts shown).")
