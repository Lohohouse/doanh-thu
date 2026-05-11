#!/usr/bin/env python3
"""
Loho House Dashboard Updater v2
Tải 8 sheets từ Google Sheets, parse data, generate Dashboard HTML.
Đầy đủ 7 hạng mục, tiêu đề tiếng Trung phồn thể.
"""

import csv
import io
import json
import os
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime

# ===== CONFIG =====
FILE_ID = "1jwPEzRMcoYBJywZkW4Vn8dKe_w5hU-zVq51zSoXGY0M"
BASE_URL = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=csv&gid="

SHEETS = {
    "shopee":      {"gid": "1202417447", "type": "processed"},
    "tiktok":      {"gid": "1316938731", "type": "processed"},
    "web":         {"gid": "1655985613", "type": "processed"},
    "lazada":      {"gid": "1217060636", "type": "processed"},
    "shopee_raw":  {"gid": "682713962",  "type": "raw"},
    "tiktok_raw":  {"gid": "1197700667", "type": "raw"},
    "web_raw":     {"gid": "1648830276", "type": "raw"},
    "lazada_raw":  {"gid": "94383236",   "type": "raw"},
}

COLS = {
    "shopee":  {"order_id": 0, "product": 14, "qty": 24, "day": 61, "month": 62, "year": 63, "revenue": 65, "fees": 66, "net": 68},
    "tiktok":  {"order_id": 0, "product": 7,  "qty": 9,  "day": 59, "month": 60, "year": 61, "revenue": 63, "fees": 64, "net": 66, "skip_rows": 2},
    "web":     {"order_id": 0, "product": 20, "qty": 19, "day": 44, "month": 45, "year": 46, "revenue": 47, "fees": 48, "net": 49},
    "lazada":  {"order_id": 4, "product": 43, "qty": None, "day": 70, "month": 71, "year": 72, "revenue": 74, "fees": 75, "net": 77},
}

RAW_COLS = {
    "shopee_raw":  {"date": 7, "product": 32, "sku": 35, "qty": 42, "price": 44},
    "tiktok_raw":  {"date": 7, "product": 23, "sku": 22, "qty": 25, "price": 31},
    "web_raw":     {"date": 17, "order_id": 1, "product": 20, "qty": 19, "price": 27},
    "lazada_raw":  {"date": 17, "order_id": 0, "product": 4, "sku": 14, "price": 55},
}

CATEGORIES = {
    "san": {
        "keywords": ["sàn nhựa", "san nhua", "sàn giả gỗ", "san gia go", "sàn dán", "san dan",
                     "combo 5m", "combo 1m", "sàn cao cấp", "san cao cap", "loho floor"],
        "sku_prefix": ["ls-", "ls_"]
    },
    "son": {
        "keywords": ["sơn tường", "son tuong", "sơn nước", "son nuoc", "sơn nội thất", "son noi that",
                     "loho-paint", "loho paint", "lọn sơn", "lon son", "paint", "sơn cao cấp"],
        "sku_prefix": ["lp-", "lp_"]
    },
    "congcu": {
        "keywords": ["băng keo", "bang keo", "cọ sơn", "co son", "cây lăn", "cay lan",
                     "khay sơn", "khay son", "công cụ", "cong cu", "dụng cụ", "dung cu",
                     "nilon", "chắn sơn", "chan son", "chắn bụi", "chan bui", "che phủ", "che phu",
                     "bông lăn", "bong lan", "bộ công cụ", "bo cong cu"],
        "sku_prefix": ["cc-", "cc_", "l-cc-", "l-cc_"]
    },
    "decor": {
        "keywords": ["decor", "trang trí", "trang tri", "câu đối", "cau doi",
                     "dây treo", "day treo", "lì xì", "li xi", "phong bao"],
        "sku_prefix": ["tt-", "tt_", "l-tt-", "l-tt_"]
    },
}


def download_csv(name, gid):
    url = BASE_URL + gid
    os.makedirs("/tmp/loho_tmp_new", exist_ok=True)
    out_path = f"/tmp/loho_tmp_new/loho_{name}.csv"
    # Reuse cache if recent (<1h old)
    if os.path.exists(out_path) and os.path.getsize(out_path) > 100:
        import time
        age = time.time() - os.path.getmtime(out_path)
        if age < 3600:
            return out_path
    result = subprocess.run(["curl", "-sL", url, "-o", out_path], capture_output=True, text=True, timeout=60)
    if result.returncode != 0:
        if os.path.exists(out_path) and os.path.getsize(out_path) > 100:
            return out_path
        print(f"ERROR downloading {name}: {result.stderr}")
        return None
    return out_path


def parse_number(s):
    if not s or not isinstance(s, str):
        return 0
    s = s.strip().replace(" ", "").replace("đ", "").replace("₫", "")
    if re.match(r'^\d\.\d{3}$', s):
        return int(s.replace(".", ""))
    if "." in s and "," not in s:
        parts = s.split(".")
        if all(len(p) == 3 for p in parts[1:]):
            s = s.replace(".", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0


def classify_product(name, sku=""):
    name_lower = (name or "").lower()
    sku_lower = (sku or "").lower()
    for cat, rules in CATEGORIES.items():
        for kw in rules["keywords"]:
            if kw in name_lower:
                return cat
        for prefix in rules["sku_prefix"]:
            if sku_lower.startswith(prefix):
                return cat
    return "other"




def canonicalize_product_name(name):
    """
    Gop cac ten SP khac nhau giua 4 san (Shopee/TikTok/Web/Lazada) ve mot ten chuan duy nhat.
    Quyet dinh gop da duoc Loho House xac nhan (2026-04-28).
    Cac SP da xac nhan TACH RIENG (giu ten goc):
      - Combo 5 cuon bang keo giay (TikTok)
      - LOHO Cong cu son cao cap, tien loi (Shopee, 62 don)
      - QUA TANG Cong cu son Loho House (TikTok, 90 don)
      - Combo 1 tay cam va 2 bong lan son (Web)
      - San may / Noi khoi nguon cam hung / Combo 4 lon son Bon mua (Web)
      - Binh Dao Dong Chu Phuc vs Chau Dao Dong Boc Nhung Do (Shopee)
    """
    if not name:
        return name
    raw = name.strip()
    n = raw.lower()

    # ===== SAN NHUA GIA GO =====
    has_san = ("sàn nhựa" in n) or ("sàn giả gỗ" in n) or ("sàn dán keo" in n)
    is_qua_tang_san = (("(quà tặng)" in n) or ("[quà tặng]" in n)) and (
        ("sàn nhựa" in n) or ("sàn giả gỗ" in n)
    )

    is_combo5 = (
        "combo 5m²" in n or "combo 5m2" in n or
        "5m² (36" in n or "5m2 (36" in n or "(36 tấm)" in n
    )
    if (has_san or is_qua_tang_san) and is_combo5:
        return "Sàn Nhựa Giả Gỗ Dán Keo COMBO 5m² (36 tấm) - 91.44x15.44cm"
    if is_qua_tang_san and "91.44" in n and "combo" not in n:
        return "Sàn Nhựa Giả Gỗ Dán Keo COMBO 5m² (36 tấm) - 91.44x15.44cm"

    is_combo1 = (
        "combo 1m²" in n or "combo 1m2" in n or
        "1m² (7" in n or "1m2 (7" in n or "(7 tấm)" in n
    )
    if has_san and is_combo1:
        return "Sàn Nhựa Giả Gỗ Dán Keo COMBO 1m² (7 tấm) - 91.44x15.44cm"

    if has_san and "91.44" in n and "combo" not in n and not is_qua_tang_san:
        return "Sàn Nhựa Giả Gỗ Dán Keo (lẻ tấm) - 91.44x15.44cm"

    # ===== SON LOT =====
    if "sơn lót" in n and "chống kiềm" in n:
        return "Sơn lót chống kiềm Loho Paint - Lon 1kg"

    # ===== SON TUONG (Lon 1kg) =====
    is_son_tuong = ("sơn tường" in n) or ("loho house wall paint" in n)
    is_mau_thu = ("mẫu thử" in n) or ("50ml" in n)
    has_1kg = "1kg" in n.replace(" ", "")
    if is_son_tuong and not is_mau_thu and has_1kg:
        if "trắng xám" in n:
            return "Sơn tường nội thất Loho-Paint tone màu Trắng Xám - Lon 1kg"
        if ("hồng" in n) or ("sweet pink" in n) or ("hoa hồng" in n):
            return "Sơn tường nội thất Loho-Paint tone màu Hồng - Lon 1kg"
        if "xanh non" in n:
            return "Sơn tường nội thất Loho-Paint tone màu Xanh Non - Lon 1kg"
        if "xanh lam" in n:
            return "Sơn tường nội thất Loho-Paint tone màu Xanh Lam - Lon 1kg"
        if "macaron" in n:
            return "Sơn tường nội thất Loho-Paint tone màu Macaron - Lon 1kg"
        if "trà sữa" in n:
            return "Sơn tường nội thất Loho-Paint tone màu Trà Sữa - Lon 1kg"
        if "cà phê" in n:
            return "Sơn tường nội thất Loho-Paint tone màu Cà Phê - Lon 1kg"
        if "nâu đất" in n:
            return "Sơn tường nội thất Loho-Paint tone màu Nâu Đất - Lon 1kg"
        if ("tone màu be" in n) or ("màu be" in n) or (" be |" in n) or (" be -" in n):
            return "Sơn tường nội thất Loho-Paint tone màu Be - Lon 1kg"
        if "trắng" in n and "xám" not in n:
            return "Sơn tường nội thất Loho-Paint tone màu Trắng - Lon 1kg"

    # ===== BO CONG CU SON 1 NGUOI (CHECK TRUOC bang keo) =====
    is_qt = ("(quà tặng)" in n) or ("[quà tặng]" in n)
    is_bo_1_nguoi = (
        ("bộ công cụ sơn" in n and "1 người" in n) or
        ("công cụ sơn cho 1 người" in n) or
        ("painting tool set for 1 person" in n)
    )
    if is_bo_1_nguoi and not is_qt:
        return "Bộ công cụ sơn 1 người Loho House"

    # ===== COMBO CAY LAN + BONG LAN (check truoc cac rule khac) =====
    if "combo" in n and "cây lăn" in n and "bông lăn" in n:
        if "1 tay cầm" not in n and "2 bông" not in n:
            return "Combo cây lăn và bông lăn sơn Loho House"

    # ===== BANG KEO NILON =====
    if "băng keo nilon" in n:
        is_qt_combo = (("(quà tặng)" in n) or ("[quà tặng]" in n)) and ("băng keo giấy" in n)
        if not is_qt_combo:
            # Khong gop neu day la 1 SP "Bo cong cu son" co liet ke "bang keo nilon" trong description
            # Ta da check is_bo_1_nguoi o tren roi nen den day chac chan KHONG phai bo cong cu
            return "Băng keo nilon chắn sơn, chắn bụi, che phủ"

    # ===== BANG KEO GIAY (cuon le) =====
    if "băng keo giấy" in n:
        is_combo5cuon = "combo 5 cuộn" in n
        is_qt = ("(quà tặng)" in n) or ("[quà tặng]" in n)
        # Check them: "1 set gom..." la dac trung cua bo cong cu, KHONG phai bang keo giay le
        is_set_listing = ("1 set gồm" in n) or ("set includes" in n) or ("1 set includes" in n)
        if not is_combo5cuon and not is_qt and not is_set_listing:
            return "Băng keo giấy chắn sơn (cuộn lẻ)"

    # ===== CO QUET SON =====
    is_co = ("cọ quét sơn" in n) or ("paint brush" in n)
    if is_co and (("3 inch" in n) or ("3 inches" in n)) and (("1.5 inch" in n) or ("1.5 inches" in n)):
        return "Cọ quét sơn cao cấp lông cước 3 inch + 1.5 inch"
    if is_co and ("cao cấp" in n) and ("(cái)" in n):
        return "Cọ quét sơn cao cấp lông cước 3 inch + 1.5 inch"

    # ===== KHAY DUNG SON 9 INCH =====
    if "khay" in n and "9 inch" in n and "đựng sơn" in n:
        return "Khay đựng sơn 9 inch Loho House"

    # ===== CAY THONG NOEL 60CM =====
    has_60cm = "60cm" in n.replace(" ", "")
    is_xmas_tree = ("cây thông" in n) or ("christmas tree" in n) or ("thông noel" in n)
    if has_60cm and is_xmas_tree:
        return "Cây thông Noel mini 60cm Loho House"

    # ===== KIT DIY MOSAIC =====
    if "kit diy" in n and (("đế lót ly" in n) or ("mosaic" in n)):
        return "Bộ Kit DIY đế lót ly mosaic Loho House"

    # ===== DUNG DICH CHONG TRON TRUOT =====
    if "chống trơn trượt" in n:
        return "Dung dịch chống trơn trượt LOHO HOUSE"

    # ===== BO CHAN GA GOI 4 MON =====
    if "chăn ga gối" in n and (("4 món" in n) or ("set 4" in n)):
        return "Bộ chăn ga gối 4 món cotton Loho House"

    # ===== THAM SOI DAY (jute) — TACH 2 mau: vintage tua rua / toi gian =====
    is_tham_jute = (("thảm" in n) or ("floor mats" in n) or ("floor mat " in n) or ("carpet" in n)) and (("sợi đay" in n) or ("jute" in n))
    is_nhat = ("nhật" in n) or ("japanese" in n)
    is_phap = ("pháp" in n) or ("french" in n)
    is_boho = "boho" in n
    if is_tham_jute and is_nhat and not is_phap and not is_boho:
        if ("vintage" in n) or ("tua rua" in n) or ("tassel" in n):
            return "Thảm tròn sợi đay vintage tua rua Nhật Bản Loho House"
        if ("tối giản" in n) or ("minimalist" in n) or ("elegant" in n):
            return "Thảm tròn sợi đay tối giản Nhật Bản Loho House"

    # ===== THAM LOT CUA KIEU PHAP =====
    if (("thảm" in n) or ("floor mat" in n) or ("door mat" in n)) and is_phap:
        return "Thảm lót cửa kiểu Pháp chống bụi"

    return raw


def read_csv(path, skip_rows=0):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        rows = list(reader)
    header_skip = 1 + skip_rows
    return rows[header_skip:] if len(rows) > header_skip else []


def safe_col(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return row[idx]


def parse_month_year(row, col_map):
    month_str = safe_col(row, col_map["month"])
    year_str = safe_col(row, col_map["year"])
    month = int(parse_number(month_str)) if month_str.strip() else 0
    year = int(parse_number(year_str)) if year_str.strip() else 0
    return month, year


def process_platform(name, csv_path):
    col_map = COLS[name]
    skip = col_map.get("skip_rows", 0)
    rows = read_csv(csv_path, skip)

    monthly = defaultdict(lambda: {
        "orders": 0, "revenue": 0, "fees": 0, "net": 0,
        "daily": defaultdict(lambda: {"revenue": 0, "orders": 0}),
        "products": defaultdict(int),
        "product_revenue": defaultdict(float),
        "order_ids": set(),
    })
    # daily_full: key = "YYYY-MM-DD" -> revenue/orders/fees/net/products/product_revenue/order_ids
    daily_full = defaultdict(lambda: {
        "revenue": 0, "orders": 0, "fees": 0, "net": 0,
        "products": defaultdict(int),
        "product_revenue": defaultdict(float),
        "order_ids": set(),
    })

    for row in rows:
        if len(row) < 5:
            continue
        month, year = parse_month_year(row, col_map)
        if month < 1 or month > 12 or year != 2026:
            continue

        key = f"T{month}"
        revenue = parse_number(safe_col(row, col_map["revenue"]))
        fees = parse_number(safe_col(row, col_map["fees"]))
        net_val = parse_number(safe_col(row, col_map["net"]))
        order_id = safe_col(row, col_map["order_id"]).strip()
        product = safe_col(row, col_map.get("product", 0)) if col_map.get("product") is not None else ""
        qty = int(parse_number(safe_col(row, col_map["qty"]))) if col_map.get("qty") is not None else 1

        if revenue <= 0:
            continue

        if name == "lazada" and net_val <= 0:
            net_val = revenue

        # Day -> full date
        day_str = safe_col(row, col_map["day"]) if col_map.get("day") is not None else ""
        full_date = None
        if day_str.strip():
            day_num = int(parse_number(day_str))
            if 1 <= day_num <= 31:
                full_date = f"{year:04d}-{month:02d}-{day_num:02d}"

        if name == "lazada":
            monthly[key]["revenue"] += revenue
            monthly[key]["fees"] += fees
            monthly[key]["net"] += net_val
            if full_date:
                daily_full[full_date]["revenue"] += revenue
                daily_full[full_date]["fees"] += fees
                daily_full[full_date]["net"] += net_val
            if order_id and order_id not in monthly[key]["order_ids"]:
                monthly[key]["order_ids"].add(order_id)
                monthly[key]["orders"] += 1
                if full_date and order_id not in daily_full[full_date]["order_ids"]:
                    daily_full[full_date]["order_ids"].add(order_id)
                    daily_full[full_date]["orders"] += 1
        elif name == "web":
            if order_id and order_id in monthly[key]["order_ids"]:
                continue
            if order_id:
                monthly[key]["order_ids"].add(order_id)
            monthly[key]["orders"] += 1
            monthly[key]["revenue"] += revenue
            monthly[key]["fees"] += fees
            monthly[key]["net"] += net_val
            if full_date:
                daily_full[full_date]["revenue"] += revenue
                daily_full[full_date]["fees"] += fees
                daily_full[full_date]["net"] += net_val
                daily_full[full_date]["orders"] += 1
                if order_id:
                    daily_full[full_date]["order_ids"].add(order_id)
        else:
            monthly[key]["orders"] += 1
            monthly[key]["revenue"] += revenue
            monthly[key]["fees"] += fees
            monthly[key]["net"] += net_val
            if full_date:
                daily_full[full_date]["revenue"] += revenue
                daily_full[full_date]["fees"] += fees
                daily_full[full_date]["net"] += net_val
                daily_full[full_date]["orders"] += 1

        if day_str.strip():
            day_num = str(int(parse_number(day_str))).zfill(2)
            monthly[key]["daily"][day_num]["revenue"] += revenue
            monthly[key]["daily"][day_num]["orders"] += 1

        if product.strip():
            canon_name = canonicalize_product_name(product.strip())
            monthly[key]["products"][canon_name] += qty if qty > 0 else 1
            monthly[key]["product_revenue"][canon_name] += revenue
            if full_date:
                daily_full[full_date]["products"][canon_name] += qty if qty > 0 else 1
                daily_full[full_date]["product_revenue"][canon_name] += revenue

    return monthly, daily_full


def process_raw_for_categories(platform, csv_path, monthly_revenue):
    raw_name = platform + "_raw"
    col_map = RAW_COLS[raw_name]
    rows = read_csv(csv_path)

    cat_totals = defaultdict(lambda: defaultdict(float))           # monthly_key -> cat -> sum
    cat_totals_daily = defaultdict(lambda: defaultdict(float))     # full_date -> cat -> sum

    for row in rows:
        if len(row) < 5:
            continue
        date_str = safe_col(row, col_map["date"]).strip()
        month = 0
        day = 0
        year = 0
        if not date_str:
            continue
        dt = None
        for fmt in ["%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y", "%d-%m-%Y"]:
            try:
                dt = datetime.strptime(date_str.strip(), fmt)
                if dt.year == 2026:
                    month = dt.month; day = dt.day; year = dt.year
                break
            except:
                dt = None
                continue
        if month == 0 and "2026" in date_str:
            try:
                dt = datetime.strptime(date_str[:15].strip(), "%d %b %Y")
                if dt.year == 2026:
                    month = dt.month; day = dt.day; year = dt.year
            except:
                pass
        if month < 1 or month > 12:
            continue

        key = f"T{month}"
        full_date = f"{year:04d}-{month:02d}-{day:02d}" if (year and day) else None
        product = safe_col(row, col_map.get("product", 0) or 0)
        sku = safe_col(row, col_map.get("sku", 0) or 0) if col_map.get("sku") is not None else ""
        price = parse_number(safe_col(row, col_map["price"]))
        if price <= 0:
            continue
        cat = classify_product(product, sku)
        cat_totals[key][cat] += price
        if full_date:
            cat_totals_daily[full_date][cat] += price

    # Normalize monthly
    result = {}
    monthly_ratios = {}
    for month_key, cats in cat_totals.items():
        actual_rev = monthly_revenue.get(month_key, 0)
        raw_total = sum(cats.values())
        if raw_total > 0 and actual_rev > 0:
            ratio = actual_rev / raw_total
            monthly_ratios[month_key] = ratio
            result[month_key] = {cat: int(val * ratio) for cat, val in cats.items()}
        else:
            monthly_ratios[month_key] = 1.0
            result[month_key] = {cat: int(val) for cat, val in cats.items()}

    # Normalize daily using its own month ratio
    result_daily = {}
    for full_date, cats in cat_totals_daily.items():
        try:
            mo = int(full_date.split("-")[1])
            ratio = monthly_ratios.get(f"T{mo}", 1.0)
        except:
            ratio = 1.0
        result_daily[full_date] = {cat: int(val * ratio) for cat, val in cats.items()}

    return result, result_daily


def enrich_products_from_raw(platform, csv_path, platforms_data, daily_data):
    """When processed sheet has empty product names (web, lazada), enrich from raw sheet.
    Also enriches daily_data[platform][full_date]['products'] with the raw product names."""
    raw_name = platform + "_raw"
    col_map = RAW_COLS[raw_name]
    rows = read_csv(csv_path)

    monthly = platforms_data[platform]
    daily_full = daily_data.get(platform, {})
    # Check if products are empty
    needs_enrichment = any(
        len(monthly[mk]["products"]) == 0
        for mk in monthly
    )
    if not needs_enrichment:
        return

    print(f"  ℹ️  Enriching {platform} products from raw sheet...")

    raw_products = defaultdict(lambda: defaultdict(lambda: {"qty": 0, "revenue": 0}))
    raw_products_daily = defaultdict(lambda: defaultdict(lambda: {"qty": 0, "revenue": 0}))

    for row in rows:
        if len(row) < 5:
            continue
        date_str = safe_col(row, col_map["date"]).strip()
        month = 0; day = 0; year = 0
        if not date_str:
            continue
        for fmt_str in ["%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y"]:
            try:
                dt = datetime.strptime(date_str.strip(), fmt_str)
                if dt.year == 2026:
                    month = dt.month; day = dt.day; year = dt.year
                break
            except:
                continue
        if month == 0 and "2026" in date_str:
            try:
                dt = datetime.strptime(date_str[:15].strip(), "%d %b %Y")
                if dt.year == 2026:
                    month = dt.month; day = dt.day; year = dt.year
            except:
                pass
        if month < 1 or month > 12:
            continue

        key = f"T{month}"
        full_date = f"{year:04d}-{month:02d}-{day:02d}" if (year and day) else None
        product = safe_col(row, col_map.get("product", 0) or 0).strip()
        if not product:
            continue
        price = parse_number(safe_col(row, col_map["price"]))
        qty_col = col_map.get("qty")
        qty = int(parse_number(safe_col(row, qty_col))) if qty_col is not None else 1
        if qty <= 0:
            qty = 1

        raw_products[key][product]["qty"] += qty
        raw_products[key][product]["revenue"] += price
        if full_date:
            raw_products_daily[full_date][product]["qty"] += qty
            raw_products_daily[full_date][product]["revenue"] += price

    # Fill in missing products (monthly)
    for mk in monthly:
        if len(monthly[mk]["products"]) == 0 and mk in raw_products:
            # Normalize revenue to match actual
            actual_rev = monthly[mk]["revenue"]
            raw_total = sum(v["revenue"] for v in raw_products[mk].values())
            ratio = actual_rev / raw_total if raw_total > 0 else 1

            for pname, pdata in raw_products[mk].items():
                canon = canonicalize_product_name(pname)
                monthly[mk]["products"][canon] = monthly[mk]["products"].get(canon, 0) + pdata["qty"]
                monthly[mk]["product_revenue"][canon] = monthly[mk]["product_revenue"].get(canon, 0) + pdata["revenue"] * ratio

    # Fill in missing products (daily_full)
    for fd in list(daily_full.keys()):
        if len(daily_full[fd].get("products", {})) == 0 and fd in raw_products_daily:
            actual_rev = daily_full[fd]["revenue"]
            raw_total = sum(v["revenue"] for v in raw_products_daily[fd].values())
            ratio = actual_rev / raw_total if raw_total > 0 else 1
            for pname, pdata in raw_products_daily[fd].items():
                canon = canonicalize_product_name(pname)
                daily_full[fd]["products"][canon] = daily_full[fd]["products"].get(canon, 0) + pdata["qty"]
                daily_full[fd]["product_revenue"][canon] = daily_full[fd]["product_revenue"].get(canon, 0) + pdata["revenue"] * ratio


def build_data_json(platforms_data, categories_data):
    result = {}
    for platform in ["shopee", "tiktok", "web", "lazada"]:
        monthly = platforms_data[platform]
        cats = categories_data.get(platform, {})
        platform_data = {}

        for month_key in sorted(monthly.keys(), key=lambda x: int(x[1:])):
            m = monthly[month_key]
            daily = {}
            for day, vals in sorted(m["daily"].items()):
                daily[day] = {"revenue": int(vals["revenue"]), "orders": vals["orders"]}

            month_cats = cats.get(month_key, {})
            categories = {
                "san": int(month_cats.get("san", 0)),
                "son": int(month_cats.get("son", 0)),
                "congcu": int(month_cats.get("congcu", 0)),
                "decor": int(month_cats.get("decor", 0)),
                "other": int(month_cats.get("other", 0)),
            }

            # Top 5 products by revenue
            products_by_rev = sorted(m["product_revenue"].items(), key=lambda x: -x[1])[:5]
            products_list = [{"name": name, "qty": m["products"].get(name, 0), "revenue": int(rev)} for name, rev in products_by_rev]

            fee_pct = round(m["fees"] / m["revenue"] * 100, 1) if m["revenue"] > 0 else 0

            platform_data[month_key] = {
                "orders": m["orders"],
                "revenue": int(m["revenue"]),
                "fees": int(m["fees"]),
                "net": int(m["net"]),
                "daily": daily,
                "categories": categories,
                "products": products_list,
                "fee_pct": fee_pct,
            }

        result[platform] = platform_data
    return result


def build_daily_json(daily_data, daily_categories_data):
    """Build full daily data: result[platform][YYYY-MM-DD] = {revenue, orders, fees, net, categories, products(top5)}"""
    result = {}
    for platform in ["shopee", "tiktok", "web", "lazada"]:
        daily_full = daily_data.get(platform, {})
        cats_daily = daily_categories_data.get(platform, {})
        platform_daily = {}
        for fd in sorted(daily_full.keys()):
            d = daily_full[fd]
            day_cats = cats_daily.get(fd, {})
            categories = {
                "san": int(day_cats.get("san", 0)),
                "son": int(day_cats.get("son", 0)),
                "congcu": int(day_cats.get("congcu", 0)),
                "decor": int(day_cats.get("decor", 0)),
                "other": int(day_cats.get("other", 0)),
            }
            products_by_rev = sorted(d["product_revenue"].items(), key=lambda x: -x[1])[:5]
            products_list = [{"name": name, "qty": d["products"].get(name, 0), "revenue": int(rev)} for name, rev in products_by_rev]
            platform_daily[fd] = {
                "orders": d["orders"],
                "revenue": int(d["revenue"]),
                "fees": int(d["fees"]),
                "net": int(d["net"]),
                "categories": categories,
                "products": products_list,
            }
        result[platform] = platform_daily
    return result


def build_products_index(daily_data):
    """Build index: product_name -> { date(YYYY-MM-DD): { platform: {qty, revenue} } }
    Use this for SP search/lookup in tab Tra Cuu.
    """
    products = {}
    for platform, daily_full in daily_data.items():
        for fd, d in daily_full.items():
            for name, qty in d["products"].items():
                rev = d["product_revenue"].get(name, 0)
                if qty <= 0 and rev <= 0:
                    continue
                if name not in products:
                    products[name] = {}
                if fd not in products[name]:
                    products[name][fd] = {}
                products[name][fd][platform] = {"qty": int(qty), "revenue": int(rev)}
    return products


def get_available_months(data):
    months = set()
    for platform in data.values():
        months.update(platform.keys())
    return sorted(months, key=lambda x: int(x[1:]))


def generate_html(data_json, daily_json, products_json, output_path):
    today = datetime.now().strftime("%d/%m/%Y %H:%M")
    months = get_available_months(data_json)
    last_month = months[-1] if months else "T1"
    data_str = json.dumps(data_json, ensure_ascii=False)
    daily_str = json.dumps(daily_json, ensure_ascii=False)
    products_str = json.dumps(products_json, ensure_ascii=False)
    trend_labels = json.dumps(months)
    # Find latest date with data for default
    all_dates = set()
    for p in daily_json.values():
        all_dates.update(p.keys())
    sorted_dates = sorted(all_dates)
    last_date = sorted_dates[-1] if sorted_dates else "2026-01-01"
    first_date = sorted_dates[0] if sorted_dates else "2026-01-01"

    month_buttons = ""
    for m in months:
        active = " active" if m == last_month else ""
        month_buttons += f'        <button class="month-btn{active}" data-month="{m}">{m}</button>\n'

    html = f'''<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard 營收報表 Báo Cáo Doanh Thu LOHO House 2026</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5; color: #333; }}
        .header {{ background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color: white; padding: 30px 20px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
        .header h1 {{ font-size: 2.5em; margin-bottom: 8px; }}
        .header p {{ font-size: 1.1em; opacity: 0.9; }}
        .controls {{ background: white; padding: 20px; display: flex; justify-content: center; gap: 10px; flex-wrap: wrap; box-shadow: 0 1px 4px rgba(0,0,0,0.05); border-bottom: 1px solid #eee; }}
        .month-btn {{ padding: 10px 20px; border: 2px solid #ddd; background: white; border-radius: 6px; cursor: pointer; font-weight: 600; transition: all 0.3s ease; }}
        .month-btn:hover {{ border-color: #e74c3c; color: #e74c3c; }}
        .month-btn.active {{ background: #e74c3c; color: white; border-color: #e74c3c; }}
        .tabs {{ display: flex; background: white; border-bottom: 2px solid #eee; padding: 0 20px; gap: 0; }}
        .tab-btn {{ padding: 15px 25px; background: none; border: none; cursor: pointer; font-weight: 600; color: #666; border-bottom: 3px solid transparent; transition: all 0.3s ease; }}
        .tab-btn:hover {{ color: #e74c3c; }}
        .tab-btn.active {{ color: #e74c3c; border-bottom-color: #e74c3c; }}
        .container {{ max-width: 1400px; margin: 20px auto; padding: 0 20px; }}
        .tab-content {{ display: none; }}
        .tab-content.active {{ display: block; }}
        .section-title {{ font-size: 1.3em; font-weight: 700; color: #1a1a2e; margin: 25px 0 15px 0; padding-bottom: 8px; border-bottom: 2px solid #e74c3c; display: inline-block; }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }}
        .kpi-card {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
        .kpi-label {{ color: #999; font-size: 0.9em; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }}
        .kpi-value {{ font-size: 1.8em; font-weight: 700; color: #1a1a2e; margin-bottom: 8px; }}
        .kpi-change {{ font-size: 0.85em; font-weight: 600; display: flex; align-items: center; gap: 5px; }}
        .chart-container {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 20px; position: relative; height: 400px; }}
        .chart-title {{ font-size: 1.1em; font-weight: 600; margin-bottom: 15px; color: #1a1a2e; }}
        .chart-wrapper {{ position: relative; height: 350px; }}
        .section-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }}
        .table-container {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow-x: auto; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        thead {{ background-color: #f8f9fa; border-bottom: 2px solid #e74c3c; }}
        th {{ padding: 12px; text-align: left; font-weight: 600; color: #1a1a2e; font-size: 0.9em; letter-spacing: 0.5px; }}
        td {{ padding: 12px; border-bottom: 1px solid #eee; }}
        tbody tr:hover {{ background-color: #f9f9f9; }}
        .badge {{ display: inline-block; padding: 4px 10px; border-radius: 4px; font-size: 0.85em; font-weight: 600; }}
        .badge.up {{ background-color: #d5f4e6; color: #27ae60; }}
        .badge.down {{ background-color: #fadbd8; color: #e74c3c; }}
        .badge.neutral {{ background-color: #eee; color: #999; }}
        th.right, td.right {{ text-align: right; }}
        @media (max-width: 768px) {{
            .kpi-grid {{ grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); }}
            .section-grid {{ grid-template-columns: 1fr; }}
            .header h1 {{ font-size: 1.8em; }}
            .tabs {{ overflow-x: auto; }}
        }}
        .lookup-bar {{ background: white; padding: 18px 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 20px; display: flex; gap: 18px; align-items: center; flex-wrap: wrap; }}
        .lookup-bar label {{ font-weight: 600; color: #1a1a2e; }}
        .lookup-bar select, .lookup-bar input {{ padding: 8px 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 0.95em; font-family: inherit; }}
        .lookup-bar input:focus, .lookup-bar select:focus {{ outline: none; border-color: #e74c3c; }}
        .lookup-bar .mode-toggle {{ display: flex; gap: 6px; }}
        .lookup-bar .mode-btn {{ padding: 7px 14px; border: 2px solid #ddd; background: white; border-radius: 6px; cursor: pointer; font-weight: 600; font-size: 0.9em; }}
        .lookup-bar .mode-btn.active {{ background: #e74c3c; color: white; border-color: #e74c3c; }}
        .lookup-bar .channel-pills {{ display: flex; gap: 6px; flex-wrap: wrap; }}
        .lookup-bar .ch-pill {{ padding: 6px 12px; border: 2px solid #ddd; background: white; border-radius: 20px; cursor: pointer; font-size: 0.85em; font-weight: 600; }}
        .lookup-bar .ch-pill.active {{ background: #1a1a2e; color: white; border-color: #1a1a2e; }}
        .lookup-bar .compare-info {{ font-size: 0.85em; color: #666; margin-left: auto; }}
        .search-bar {{ position: relative; margin-bottom: 20px; }}
        .search-bar input {{ width: 100%; padding: 12px 16px; border: 2px solid #ddd; border-radius: 8px; font-size: 15px; font-family: inherit; box-sizing: border-box; }}
        .search-bar input:focus {{ outline: none; border-color: #e74c3c; }}
        .search-results {{ position: absolute; top: calc(100% + 2px); left: 0; right: 0; background: white; border: 1px solid #ddd; border-radius: 8px; max-height: 320px; overflow-y: auto; z-index: 100; display: none; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }}
        .search-results.show {{ display: block; }}
        .search-result-item {{ padding: 10px 16px; cursor: pointer; border-bottom: 1px solid #f0f0f0; font-size: 14px; transition: background 0.15s; }}
        .search-result-item:hover {{ background: #fef3f0; }}
        .search-result-item:last-child {{ border-bottom: none; }}
        .search-result-item mark {{ background: #fff3cd; padding: 0 2px; border-radius: 2px; }}
        .search-result-item .meta {{ font-size: 12px; color: #999; margin-top: 3px; }}
        #lookup-sp-detail {{ background: #fafafa; padding: 20px; border-radius: 8px; margin-bottom: 20px; }}
        #lookup-sp-name {{ font-size: 1.05em; color: #1a1a2e; margin-bottom: 18px; padding: 12px; background: #fff; border-radius: 6px; border-left: 4px solid #e74c3c; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>LOHO House 營收報表 Báo Cáo Doanh Thu 2026</h1>
        <p>更新時間 Cập nhật: {today}</p>
    </div>

    <div class="controls">
{month_buttons}    </div>

    <div class="tabs">
        <button class="tab-btn active" data-tab="tong-quan">總覽 Tổng Quan</button>
        <button class="tab-btn" data-tab="shopee">Shopee</button>
        <button class="tab-btn" data-tab="tiktok">TikTok</button>
        <button class="tab-btn" data-tab="lazada">Lazada</button>
        <button class="tab-btn" data-tab="website">Website</button>
        <button class="tab-btn" data-tab="lookup">查詢 Tra Cứu Theo Ngày</button>
    </div>

    <div class="container">
        <!-- ===== 總覽 TAB ===== -->
        <div id="tong-quan" class="tab-content active">
            <div class="section-title">① 關鍵績效指標總覽 KPI Tổng Quan</div>
            <div class="kpi-grid" id="tongQuan-kpi"></div>

            <div class="section-title">② 每日營收圖表 Biểu Đồ Doanh Thu Theo Ngày</div>
            <div class="chart-container"><div class="chart-title">每日營收趨勢 Xu hướng doanh thu hàng ngày</div><div class="chart-wrapper"><canvas id="tongQuan-daily-chart"></canvas></div></div>

            <div class="section-title">③ 各月份營收成長比較 So Sánh Tăng Trưởng Doanh Thu &amp; 各通路比較 Các Kênh</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">各月份營收趨勢 Xu hướng doanh thu các tháng</div><div class="chart-wrapper"><canvas id="tongQuan-trend-chart"></canvas></div></div>
                <div class="chart-container"><div class="chart-title">各通路營收比較 So sánh kênh (本月 vs 上月 Tháng này vs Tháng trước)</div><div class="chart-wrapper"><canvas id="tongQuan-channel-chart"></canvas></div></div>
            </div>
            <div class="table-container">
                <div class="chart-title">各月份營收成長明細 Chi tiết tăng trưởng doanh thu từng tháng</div>
                <table><thead><tr><th>月份 Tháng</th><th class="right">營收 Doanh Thu</th><th class="right">訂單數 Đơn Hàng</th><th class="right">客單價 AOV</th><th>環比成長 Tăng Trưởng</th></tr></thead><tbody id="tongQuan-mom-table"></tbody></table>
            </div>

            <div class="section-title">④ 產品類別營收 Doanh Thu Theo Danh Mục Sản Phẩm</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">產品類別營收分佈 Phân bổ doanh thu danh mục</div><div class="chart-wrapper"><canvas id="tongQuan-category-chart"></canvas></div></div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">各類別營收明細 Chi tiết doanh thu từng danh mục</div>
                    <table><thead><tr><th>類別 Danh Mục</th><th class="right">本月營收 Tháng Này</th><th class="right">上月營收 Tháng Trước</th><th>變化 Thay Đổi</th></tr></thead><tbody id="tongQuan-category-table"></tbody></table>
                </div>
            </div>

            <div class="section-title">⑤ 平台費用分析 Chi Phí Sàn &amp; Tỷ Lệ Phí</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">各通路平台費用 Phí sàn các kênh</div><div class="chart-wrapper"><canvas id="tongQuan-fees-chart"></canvas></div></div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">平台費用及費率 Phí sàn và tỷ lệ %</div>
                    <table><thead><tr><th>通路 Kênh</th><th class="right">營收 Doanh Thu</th><th class="right">平台費用 Phí Sàn</th><th class="right">費率 Tỷ Lệ %</th><th class="right">淨收入 DT Ròng</th></tr></thead><tbody id="tongQuan-fees-table"></tbody></table>
                </div>
            </div>

            <div class="section-title">⑥ 各月份營收環比比較 So Sánh Doanh Thu Tháng Với Tháng Trước</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">各通路月度營收對比 So sánh doanh thu kênh theo tháng</div><div class="chart-wrapper"><canvas id="tongQuan-mom-channel-chart"></canvas></div></div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">各通路本月 vs 上月 Các kênh: Tháng này vs Tháng trước</div>
                    <table><thead><tr><th>通路 Kênh</th><th class="right">本月營收 Tháng Này</th><th class="right">上月營收 Tháng Trước</th><th>變化 Thay Đổi</th></tr></thead><tbody id="tongQuan-mom-channel-table"></tbody></table>
                </div>
            </div>

            <div class="section-title">⑦ 暢銷產品類別 Danh Mục SP Bán Chạy Theo Doanh Thu</div>
            <div class="table-container">
                <div class="chart-title">暢銷產品 Top 5 Top 5 sản phẩm bán chạy (so sánh tháng trước)</div>
                <table><thead><tr><th>#</th><th>產品名稱 Tên SP</th><th class="right">數量 SL</th><th class="right">營收 Doanh Thu</th><th>通路 Kênh</th></tr></thead><tbody id="tongQuan-products-table"></tbody></table>
            </div>
        </div>

        <!-- ===== TRA CUU THEO NGAY TAB ===== -->
        <div id="lookup" class="tab-content">
            <div class="lookup-bar">
                <div class="mode-toggle">
                    <button class="mode-btn active" data-mode="single" data-target="main">1 ngày</button>
                    <button class="mode-btn" data-mode="range" data-target="main">Khoảng ngày</button>
                </div>
                <div id="lookup-single-controls">
                    <label>Chọn ngày:</label>
                    <input type="date" id="lookup-date" value="{last_date}" min="{first_date}" max="{last_date}">
                </div>
                <div id="lookup-range-controls" style="display:none;">
                    <label>Từ:</label>
                    <input type="date" id="lookup-start" value="{first_date}" min="{first_date}" max="{last_date}">
                    <label>Đến:</label>
                    <input type="date" id="lookup-end" value="{last_date}" min="{first_date}" max="{last_date}">
                </div>
                <div class="channel-pills">
                    <button class="ch-pill active" data-channel="all" data-target="main">Tất cả</button>
                    <button class="ch-pill" data-channel="shopee" data-target="main">Shopee</button>
                    <button class="ch-pill" data-channel="tiktok" data-target="main">TikTok</button>
                    <button class="ch-pill" data-channel="web" data-target="main">Web</button>
                    <button class="ch-pill" data-channel="lazada" data-target="main">Lazada</button>
                </div>
                <div class="compare-info" id="lookup-compare-info"></div>
            </div>

            <div class="section-title">① 關鍵績效指標 KPI Tra Cứu</div>
            <div class="kpi-grid" id="lookup-kpi"></div>

            <div class="section-title">② 每日營收圖表 Biểu Đồ Doanh Thu Theo Ngày</div>
            <div class="chart-container"><div class="chart-title">營收趨勢 Xu hướng doanh thu theo ngày trong khoảng đã chọn</div><div class="chart-wrapper"><canvas id="lookup-daily-chart"></canvas></div></div>

            <div class="section-title">③ 對比成長 So Sánh Tăng Trưởng (vs Hôm Trước &amp; Cùng Kỳ Tuần Trước)</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">營收對比 So sánh doanh thu (3 kỳ)</div><div class="chart-wrapper"><canvas id="lookup-trend-chart"></canvas></div></div>
                <div class="chart-container"><div class="chart-title">各通路營收比較 So sánh kênh (kỳ này)</div><div class="chart-wrapper"><canvas id="lookup-channel-chart"></canvas></div></div>
            </div>
            <div class="table-container">
                <div class="chart-title">對比明細 Chi tiết so sánh kỳ này / hôm trước / cùng kỳ tuần trước</div>
                <table><thead><tr><th>Kỳ So Sánh</th><th class="right">營收 Doanh Thu</th><th class="right">訂單 Đơn Hàng</th><th class="right">客單價 AOV</th><th>變化 vs Kỳ Này</th></tr></thead><tbody id="lookup-mom-table"></tbody></table>
            </div>

            <div class="section-title">④ 產品類別營收 Doanh Thu Theo Danh Mục Sản Phẩm</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">產品類別營收分佈 Phân bổ doanh thu danh mục</div><div class="chart-wrapper"><canvas id="lookup-category-chart"></canvas></div></div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">各類別營收明細 Chi tiết doanh thu từng danh mục</div>
                    <table><thead><tr><th>類別 Danh Mục</th><th class="right">本期營收 Kỳ Này</th><th class="right">上期 Hôm Trước</th><th>變化 Thay Đổi</th></tr></thead><tbody id="lookup-category-table"></tbody></table>
                </div>
            </div>

            <div class="section-title">⑤ 平台費用分析 Chi Phí Sàn &amp; Tỷ Lệ Phí</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">各通路平台費用 Phí sàn các kênh (kỳ này)</div><div class="chart-wrapper"><canvas id="lookup-fees-chart"></canvas></div></div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">平台費用及費率 Phí sàn và tỷ lệ %</div>
                    <table><thead><tr><th>通路 Kênh</th><th class="right">營收 Doanh Thu</th><th class="right">平台費用 Phí Sàn</th><th class="right">費率 Tỷ Lệ %</th><th class="right">淨收入 DT Ròng</th></tr></thead><tbody id="lookup-fees-table"></tbody></table>
                </div>
            </div>

            <div class="section-title">⑥ 各通路對比 So Sánh Doanh Thu Các Kênh</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">各通路對比 Kỳ này vs Hôm trước vs Cùng kỳ tuần trước</div><div class="chart-wrapper"><canvas id="lookup-mom-channel-chart"></canvas></div></div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">各通路詳細 Chi tiết các kênh (so sánh hôm trước)</div>
                    <table><thead><tr><th>通路 Kênh</th><th class="right">本期 Kỳ Này</th><th class="right">上期 Hôm Trước</th><th class="right">同期週前 Tuần Trước</th><th>變化 vs Hôm Trước</th></tr></thead><tbody id="lookup-mom-channel-table"></tbody></table>
                </div>
            </div>

            <div class="section-title">⑦ 暢銷產品 SP Bán Chạy Theo Doanh Thu</div>
            <div class="table-container">
                <div class="chart-title">Top 5 暢銷產品 Sản phẩm bán chạy trong kỳ đã chọn</div>
                <table><thead><tr><th>#</th><th>產品名稱 Tên SP</th><th class="right">數量 SL</th><th class="right">營收 Doanh Thu</th></tr></thead><tbody id="lookup-products-table"></tbody></table>
            </div>

            <div class="section-title">🔍 Tra Cứu Chi Tiết 1 Sản Phẩm Theo Ngày &amp; Theo Kênh</div>
            <div class="lookup-bar">
                <div class="mode-toggle">
                    <button class="mode-btn active" data-mode="single" data-target="sp">1 ngày</button>
                    <button class="mode-btn" data-mode="range" data-target="sp">Khoảng ngày</button>
                </div>
                <div id="sp-single-controls">
                    <label>Chọn ngày:</label>
                    <input type="date" id="sp-date" value="{last_date}" min="{first_date}" max="{last_date}">
                </div>
                <div id="sp-range-controls" style="display:none;">
                    <label>Từ:</label>
                    <input type="date" id="sp-start" value="{first_date}" min="{first_date}" max="{last_date}">
                    <label>Đến:</label>
                    <input type="date" id="sp-end" value="{last_date}" min="{first_date}" max="{last_date}">
                </div>
                <div class="channel-pills">
                    <button class="ch-pill active" data-channel="all" data-target="sp">Tất cả</button>
                    <button class="ch-pill" data-channel="shopee" data-target="sp">Shopee</button>
                    <button class="ch-pill" data-channel="tiktok" data-target="sp">TikTok</button>
                    <button class="ch-pill" data-channel="web" data-target="sp">Web</button>
                    <button class="ch-pill" data-channel="lazada" data-target="sp">Lazada</button>
                </div>
                <div class="compare-info" id="sp-compare-info"></div>
            </div>
            <div class="search-bar">
                <input type="text" id="lookup-sp-search" placeholder="🔍 Gõ tên SP để tìm (vd: 'Sàn COMBO 5m' hoặc 'Sơn Be' hoặc 'Băng keo')..." autocomplete="off">
                <div id="lookup-sp-results" class="search-results"></div>
            </div>

            <div id="lookup-sp-detail" style="display:none;">
                <div id="lookup-sp-name"></div>
                <div class="kpi-grid" id="lookup-sp-kpi"></div>

                <div class="section-grid">
                    <div class="table-container" style="height:auto;">
                        <div class="chart-title">📊 Phân bổ theo Kênh</div>
                        <table><thead><tr><th>Kênh</th><th class="right">SL</th><th class="right">Doanh Thu</th><th class="right">% Tổng</th></tr></thead><tbody id="lookup-sp-channel-table"></tbody></table>
                    </div>
                    <div class="table-container" style="height:auto;">
                        <div class="chart-title">📅 Chi tiết theo Ngày</div>
                        <table><thead><tr><th>Ngày</th><th class="right">SL</th><th class="right">Doanh Thu</th><th>Kênh chính</th></tr></thead><tbody id="lookup-sp-date-table"></tbody></table>
                    </div>
                </div>
            </div>
        </div>

        <!-- ===== PLATFORM TABS ===== -->''' + '''
''' + _generate_platform_tabs() + f'''
    </div>

    <script>
const D = {data_str};
const allMonths = {trend_labels};
const categoryNames={{san:"地板",son:"油漆",congcu:"工具",decor:"裝飾",other:"其他"}};
const categoryNamesVi={{san:"Sàn nhựa",son:"Sơn tường",congcu:"Công cụ",decor:"Decor",other:"Khác"}};
const chartColors=["#e74c3c","#3498db","#2ecc71","#f39c12","#9b59b6"];
const platformNames={{shopee:"Shopee",tiktok:"TikTok",web:"Website",lazada:"Lazada"}};
const tabToKey={{shopee:"shopee",tiktok:"tiktok",lazada:"lazada",website:"web"}};
function resolveKey(tid){{return tabToKey[tid]||tid;}}
let currentMonth="{last_month}",charts={{}};
const DD={daily_str};
const PRODUCTS={products_str};
const lookupFirstDate="{first_date}";
const lookupLastDate="{last_date}";
let lookupMode="single";
let lookupChannel="all";

function fmt(n){{if(n>=1e9)return (n/1e9).toFixed(2)+"tỷ";if(n>=1e6)return (n/1e6).toFixed(1)+"tr";return n.toString().replace(/\\B(?=(\\d{{3}})+(?!\\d))/g,".");}}
function fmtFull(n){{return n.toString().replace(/\\B(?=(\\d{{3}})+(?!\\d))/g,".");}}
function chg(c,p){{if(!p)return{{pct:"N/A",arrow:"—",color:"#999"}};const pct=((c-p)/p*100).toFixed(1);return{{pct,arrow:pct>=0?"↑":"↓",color:pct>=0?"#27ae60":"#e74c3c"}};}}
function getPrev(){{const idx=allMonths.indexOf(currentMonth);return idx>0?allMonths[idx-1]:null;}}
function getD(p,m){{
    if(!m||!D[p]||!D[p][m]) return {{orders:0,revenue:0,fees:0,net:0,daily:{{}},categories:{{san:0,son:0,congcu:0,decor:0,other:0}},products:[],fee_pct:0}};
    return D[p][m];
}}
function badgeHtml(cur,prev){{
    if(!prev)return '<span class="badge neutral">—</span>';
    const delta=((cur-prev)/prev*100).toFixed(1);
    return delta>=0?`<span class="badge up">↑ ${{delta}}%</span>`:`<span class="badge down">↓ ${{Math.abs(delta).toFixed(1)}}%</span>`;
}}

/* ① KPI */
function renderKpi(platforms,elId){{
    const pm=getPrev();
    let tr=0,to=0,tf=0,tn=0,tr2=0,to2=0,tf2=0,tn2=0;
    platforms.forEach(p=>{{const d=getD(p,currentMonth),prev=getD(p,pm);tr+=d.revenue;to+=d.orders;tf+=d.fees;tn+=d.net;tr2+=prev.revenue;to2+=prev.orders;tf2+=prev.fees;tn2+=prev.net;}});
    const aov=to?Math.round(tr/to):0,aov2=to2?Math.round(tr2/to2):0;
    const k=[
        {{l:"營收 Doanh Thu",v:tr,c:chg(tr,tr2)}},
        {{l:"訂單 Đơn Hàng",v:to,c:chg(to,to2)}},
        {{l:"客單價 AOV",v:aov,c:chg(aov,aov2)}},
        {{l:"平台費 Phí Sàn",v:tf,c:chg(tf,tf2)}},
        {{l:"淨收入 DT Ròng",v:tn,c:chg(tn,tn2)}}
    ];
    let h="";k.forEach(x=>{{h+=`<div class="kpi-card"><div class="kpi-label">${{x.l}}</div><div class="kpi-value">${{fmt(x.v)}}</div><div class="kpi-change" style="color:${{x.c.color}}">${{x.c.arrow}} ${{x.c.pct}}%</div></div>`;}});
    document.getElementById(elId).innerHTML=h;
}}

/* ② Daily chart */
function renderDaily(platforms,canvasId){{
    const ad={{}};
    platforms.forEach(p=>{{const d=getD(p,currentMonth);Object.keys(d.daily).forEach(y=>{{if(!ad[y])ad[y]=0;ad[y]+=d.daily[y].revenue;}});}});
    const ds=Object.keys(ad).sort((a,b)=>parseInt(a)-parseInt(b)),dvs=ds.map(x=>ad[x]);
    if(charts[canvasId])charts[canvasId].destroy();
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"line",data:{{labels:ds,datasets:[{{label:"營收",data:dvs,borderColor:"#e74c3c",backgroundColor:"rgba(231,76,60,0.1)",fill:true,tension:.4}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
}}

/* ③ Trend + channel comparison */
function renderTrend(platforms,canvasId){{
    if(platforms.length===1){{
        const p=platforms[0];
        const data=allMonths.map(m=>getD(p,m).revenue);
        if(charts[canvasId])charts[canvasId].destroy();
        charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels:allMonths,datasets:[{{label:"營收",data:data,backgroundColor:"#e74c3c"}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
    }} else {{
        const datasets=["shopee","tiktok","web","lazada"].map((p,i)=>({{label:platformNames[p],data:allMonths.map(m=>getD(p,m).revenue),backgroundColor:chartColors[i]}}));
        if(charts[canvasId])charts[canvasId].destroy();
        charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels:allMonths,datasets}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
    }}
}}

function renderChannelCompare(canvasId){{
    const pm=getPrev();
    const labels=["Shopee","TikTok","Web","Lazada"];
    const cur=["shopee","tiktok","web","lazada"].map(p=>getD(p,currentMonth).revenue);
    const prev=["shopee","tiktok","web","lazada"].map(p=>getD(p,pm).revenue);
    if(charts[canvasId])charts[canvasId].destroy();
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels,datasets:[{{label:"上月",data:prev,backgroundColor:"#95a5a6"}},{{label:"本月",data:cur,backgroundColor:"#e74c3c"}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
}}

function renderMomTable(platforms,elId){{
    let h="";
    allMonths.forEach((m,i)=>{{
        let rev=0,ord=0;
        platforms.forEach(p=>{{const d=getD(p,m);rev+=d.revenue;ord+=d.orders;}});
        const aov=ord?Math.round(rev/ord):0;
        let prevRev=0;
        if(i>0){{platforms.forEach(p=>{{prevRev+=getD(p,allMonths[i-1]).revenue;}});}}
        const badge=i>0?badgeHtml(rev,prevRev):'<span class="badge neutral">—</span>';
        h+=`<tr><td>${{m}}</td><td class="right">${{fmtFull(Math.round(rev))}}</td><td class="right">${{fmtFull(ord)}}</td><td class="right">${{fmtFull(aov)}}</td><td>${{badge}}</td></tr>`;
    }});
    document.getElementById(elId).innerHTML=h;
}}

/* ④ Category */
function renderCategoryChart(platforms,canvasId){{
    const ac={{san:0,son:0,congcu:0,decor:0,other:0}};
    platforms.forEach(p=>{{const d=getD(p,currentMonth);Object.keys(ac).forEach(y=>{{ac[y]+=d.categories[y];}});}});
    if(charts[canvasId])charts[canvasId].destroy();
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"doughnut",data:{{labels:Object.keys(ac).map(x=>categoryNames[x]+" "+categoryNamesVi[x]),datasets:[{{data:Object.values(ac),backgroundColor:chartColors}}]}},options:{{responsive:true,maintainAspectRatio:false}}}});
}}

function renderCategoryTable(platforms,elId){{
    const pm=getPrev();
    const cats=["san","son","congcu","decor","other"];
    const cur={{}},prev={{}};
    cats.forEach(c=>{{cur[c]=0;prev[c]=0;}});
    platforms.forEach(p=>{{
        const d=getD(p,currentMonth),d2=getD(p,pm);
        cats.forEach(c=>{{cur[c]+=d.categories[c];prev[c]+=d2.categories[c];}});
    }});
    // Sort by revenue descending
    const sorted=[...cats].sort((a,b)=>cur[b]-cur[a]);
    let h="";
    sorted.forEach(c=>{{
        h+=`<tr><td>${{categoryNames[c]}} ${{categoryNamesVi[c]}}</td><td class="right">${{fmtFull(cur[c])}}</td><td class="right">${{fmtFull(prev[c])}}</td><td>${{badgeHtml(cur[c],prev[c])}}</td></tr>`;
    }});
    document.getElementById(elId).innerHTML=h;
}}

/* ⑤ Fees */
function renderFeesChart(canvasId){{
    const labels=["Shopee","TikTok","Web","Lazada"];
    const fees=["shopee","tiktok","web","lazada"].map(p=>getD(p,currentMonth).fees);
    if(charts[canvasId])charts[canvasId].destroy();
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels,datasets:[{{label:"平台費用",data:fees,backgroundColor:chartColors}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
}}

function renderFeesTable(elId){{
    let h="",totalRev=0,totalFees=0,totalNet=0;
    ["shopee","tiktok","web","lazada"].forEach(p=>{{
        const d=getD(p,currentMonth);
        const pct=d.revenue?((d.fees/d.revenue)*100).toFixed(1):"-";
        h+=`<tr><td>${{platformNames[p]}}</td><td class="right">${{fmtFull(d.revenue)}}</td><td class="right">${{fmtFull(d.fees)}}</td><td class="right">${{pct}}%</td><td class="right">${{fmtFull(d.net)}}</td></tr>`;
        totalRev+=d.revenue;totalFees+=d.fees;totalNet+=d.net;
    }});
    const totalPct=totalRev?((totalFees/totalRev)*100).toFixed(1):"-";
    h+=`<tr style="font-weight:700;border-top:2px solid #333"><td>合計 Tổng Cộng</td><td class="right">${{fmtFull(totalRev)}}</td><td class="right">${{fmtFull(totalFees)}}</td><td class="right">${{totalPct}}%</td><td class="right">${{fmtFull(totalNet)}}</td></tr>`;
    document.getElementById(elId).innerHTML=h;
}}

/* ⑥ MoM channel */
function renderMomChannelChart(canvasId){{
    const pm=getPrev();
    const platforms=["shopee","tiktok","web","lazada"];
    const labels=platforms.map(p=>platformNames[p]);
    const cur=platforms.map(p=>getD(p,currentMonth).revenue);
    const prev=platforms.map(p=>getD(p,pm).revenue);
    if(charts[canvasId])charts[canvasId].destroy();
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels,datasets:[{{label:"上月",data:prev,backgroundColor:"#95a5a6"}},{{label:"本月",data:cur,backgroundColor:"#e74c3c"}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
}}

function renderMomChannelTable(elId){{
    const pm=getPrev();
    let h="";
    ["shopee","tiktok","web","lazada"].forEach(p=>{{
        const d=getD(p,currentMonth),prev=getD(p,pm);
        h+=`<tr><td>${{platformNames[p]}}</td><td class="right">${{fmtFull(d.revenue)}}</td><td class="right">${{fmtFull(prev.revenue)}}</td><td>${{badgeHtml(d.revenue,prev.revenue)}}</td></tr>`;
    }});
    document.getElementById(elId).innerHTML=h;
}}

/* ⑦ Products */
function renderProducts(platforms,elId,showChannel){{
    const pm=getPrev();
    const ap={{}};
    platforms.forEach(p=>{{
        const d=getD(p,currentMonth);
        d.products.forEach(x=>{{
            const key=x.name;
            if(!ap[key])ap[key]={{qty:0,revenue:0,channel:platformNames[p]}};
            ap[key].qty+=x.qty;ap[key].revenue+=(x.revenue||0);
            if(platforms.length>1)ap[key].channel="Mixed";
        }});
    }});
    const sorted=Object.entries(ap).sort((a,b)=>b[1].revenue-a[1].revenue).slice(0,5);
    let h="";
    sorted.forEach(([name,v],i)=>{{
        h+=`<tr><td>${{i+1}}</td><td>${{name}}</td><td class="right">${{fmtFull(v.qty)}}</td><td class="right">${{fmtFull(v.revenue)}}</td>`;
        if(showChannel)h+=`<td>${{v.channel}}</td>`;
        h+="</tr>";
    }});
    document.getElementById(elId).innerHTML=h;
}}

/* Platform-specific order trend */
function renderOrdersTrend(platform,canvasId){{
    const data=allMonths.map(m=>getD(platform,m).orders);
    if(charts[canvasId])charts[canvasId].destroy();
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels:allMonths,datasets:[{{label:"訂單",data:data,backgroundColor:"#3498db"}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
}}

/* Platform-specific fee for single platform */
function renderPlatformFees(platform,chartId,tableId){{
    const feesData=allMonths.map(m=>getD(platform,m).fees);
    if(charts[chartId])charts[chartId].destroy();
    charts[chartId]=new Chart(document.getElementById(chartId),{{type:"bar",data:{{labels:allMonths,datasets:[{{label:"平台費用",data:feesData,backgroundColor:"#f39c12"}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
    let h="";
    allMonths.forEach(m=>{{
        const d=getD(platform,m);
        const pct=d.revenue?((d.fees/d.revenue)*100).toFixed(1):"-";
        h+=`<tr><td>${{m}}</td><td class="right">${{fmtFull(d.revenue)}}</td><td class="right">${{fmtFull(d.fees)}}</td><td class="right">${{pct}}%</td><td class="right">${{fmtFull(d.net)}}</td></tr>`;
    }});
    document.getElementById(tableId).innerHTML=h;
}}

/* ===== LOOKUP HELPERS ===== */
function ymdAdd(ymd, days){{
    const d=new Date(ymd+"T00:00:00");
    d.setDate(d.getDate()+days);
    const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,"0"),dd=String(d.getDate()).padStart(2,"0");
    return `${{y}}-${{m}}-${{dd}}`;
}}
function ymdDiff(a,b){{
    return Math.round((new Date(b+"T00:00:00")-new Date(a+"T00:00:00"))/86400000);
}}
function fmtDate(ymd){{ const p=ymd.split("-"); return `${{p[2]}}/${{p[1]}}/${{p[0]}}`; }}

function getLookupRange(){{
    if(lookupMode==="single"){{
        const v=document.getElementById("lookup-date").value||lookupLastDate;
        return [v,v];
    }} else {{
        let s=document.getElementById("lookup-start").value||lookupFirstDate;
        let e=document.getElementById("lookup-end").value||lookupLastDate;
        if(s>e){{const t=s;s=e;e=t;}}
        return [s,e];
    }}
}}
function getLookupPlatforms(){{
    return lookupChannel==="all"?["shopee","tiktok","web","lazada"]:[lookupChannel];
}}

/* Aggregate ranges of daily data */
function aggLookup(start,end,platforms){{
    const result={{revenue:0,orders:0,fees:0,net:0,daily:{{}},categories:{{san:0,son:0,congcu:0,decor:0,other:0}},products:[],byChannel:{{shopee:0,tiktok:0,web:0,lazada:0}},feesByChannel:{{shopee:0,tiktok:0,web:0,lazada:0}},netByChannel:{{shopee:0,tiktok:0,web:0,lazada:0}},ordersByChannel:{{shopee:0,tiktok:0,web:0,lazada:0}}}};
    const productMap={{}};
    platforms.forEach(p=>{{
        if(!DD[p])return;
        Object.keys(DD[p]).forEach(d=>{{
            if(d>=start&&d<=end){{
                const day=DD[p][d];
                result.revenue+=day.revenue;
                result.orders+=day.orders;
                result.fees+=day.fees;
                result.net+=day.net;
                result.byChannel[p]=(result.byChannel[p]||0)+day.revenue;
                result.feesByChannel[p]=(result.feesByChannel[p]||0)+day.fees;
                result.netByChannel[p]=(result.netByChannel[p]||0)+day.net;
                result.ordersByChannel[p]=(result.ordersByChannel[p]||0)+day.orders;
                Object.keys(result.categories).forEach(k=>result.categories[k]+=(day.categories[k]||0));
                if(!result.daily[d])result.daily[d]={{revenue:0,orders:0}};
                result.daily[d].revenue+=day.revenue;
                result.daily[d].orders+=day.orders;
                (day.products||[]).forEach(prod=>{{
                    if(!productMap[prod.name])productMap[prod.name]={{qty:0,revenue:0}};
                    productMap[prod.name].qty+=prod.qty;
                    productMap[prod.name].revenue+=prod.revenue;
                }});
            }}
        }});
    }});
    result.products=Object.entries(productMap).map(([n,v])=>({{name:n,qty:v.qty,revenue:v.revenue}})).sort((a,b)=>b.revenue-a.revenue).slice(0,5);
    return result;
}}

function getPrevDayRange(s,e){{
    const len=ymdDiff(s,e);
    const ne=ymdAdd(s,-1);
    const ns=ymdAdd(ne,-len);
    return [ns,ne];
}}
function getPrevWeekRange(s,e){{
    return [ymdAdd(s,-7),ymdAdd(e,-7)];
}}

function lookup(){{
    const [s,e]=getLookupRange();
    const platforms=getLookupPlatforms();
    const cur=aggLookup(s,e,platforms);
    const [ps,pe]=getPrevDayRange(s,e);
    const prev=aggLookup(ps,pe,platforms);
    const [ws,we]=getPrevWeekRange(s,e);
    const prevWeek=aggLookup(ws,we,platforms);

    // Update info bar
    const infoEl=document.getElementById("lookup-compare-info");
    if(infoEl){{
        const same=(s===e);
        infoEl.innerHTML = same
            ? `So sánh: <b>${{fmtDate(s)}}</b> vs <b>${{fmtDate(ps)}}</b> (hôm trước) &amp; <b>${{fmtDate(ws)}}</b> (tuần trước)`
            : `Kỳ này: <b>${{fmtDate(s)}} → ${{fmtDate(e)}}</b> (${{ymdDiff(s,e)+1}} ngày) | Hôm trước: ${{fmtDate(ps)}} → ${{fmtDate(pe)}} | Tuần trước: ${{fmtDate(ws)}} → ${{fmtDate(we)}}`;
    }}

    /* (1) KPI */
    const aov=cur.orders?Math.round(cur.revenue/cur.orders):0;
    const aov2=prev.orders?Math.round(prev.revenue/prev.orders):0;
    const k=[
        {{l:"營收 Doanh Thu",v:cur.revenue,c:chg(cur.revenue,prev.revenue)}},
        {{l:"訂單 Đơn Hàng",v:cur.orders,c:chg(cur.orders,prev.orders)}},
        {{l:"客單價 AOV",v:aov,c:chg(aov,aov2)}},
        {{l:"平台費 Phí Sàn",v:cur.fees,c:chg(cur.fees,prev.fees)}},
        {{l:"淨收入 DT Ròng",v:cur.net,c:chg(cur.net,prev.net)}}
    ];
    let h="";k.forEach(x=>{{h+=`<div class="kpi-card"><div class="kpi-label">${{x.l}}</div><div class="kpi-value">${{fmt(x.v)}}</div><div class="kpi-change" style="color:${{x.c.color}}">${{x.c.arrow}} ${{x.c.pct}}%</div></div>`;}});
    document.getElementById("lookup-kpi").innerHTML=h;

    /* (2) Daily chart */
    const days=Object.keys(cur.daily).sort();
    const dvs=days.map(d=>cur.daily[d].revenue);
    const dlabels=days.map(d=>fmtDate(d));
    if(charts["lookup-daily-chart"])charts["lookup-daily-chart"].destroy();
    charts["lookup-daily-chart"]=new Chart(document.getElementById("lookup-daily-chart"),{{type:"line",data:{{labels:dlabels,datasets:[{{label:"營收",data:dvs,borderColor:"#e74c3c",backgroundColor:"rgba(231,76,60,0.1)",fill:true,tension:.4}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});

    /* (3) Trend compare */
    if(charts["lookup-trend-chart"])charts["lookup-trend-chart"].destroy();
    charts["lookup-trend-chart"]=new Chart(document.getElementById("lookup-trend-chart"),{{type:"bar",data:{{labels:["週前","昨日","本期"],datasets:[{{label:"營收",data:[prevWeek.revenue,prev.revenue,cur.revenue],backgroundColor:["#9b59b6","#95a5a6","#e74c3c"]}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});

    /* (3b) Channel compare */
    const chLabels=["Shopee","TikTok","Web","Lazada"];
    const chCur=["shopee","tiktok","web","lazada"].map(p=>cur.byChannel[p]||0);
    if(charts["lookup-channel-chart"])charts["lookup-channel-chart"].destroy();
    charts["lookup-channel-chart"]=new Chart(document.getElementById("lookup-channel-chart"),{{type:"bar",data:{{labels:chLabels,datasets:[{{label:"營收",data:chCur,backgroundColor:chartColors}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});

    /* (3c) MoM table - 3 hang so sanh */
    const aov3=prevWeek.orders?Math.round(prevWeek.revenue/prevWeek.orders):0;
    const rows=[
        {{label:"本期 Kỳ Này",rev:cur.revenue,ord:cur.orders,aov,vsCur:'<span class="badge neutral">—</span>'}},
        {{label:"昨日 Hôm Trước",rev:prev.revenue,ord:prev.orders,aov:aov2,vsCur:badgeHtml(prev.revenue,cur.revenue)}},
        {{label:"週前 Tuần Trước",rev:prevWeek.revenue,ord:prevWeek.orders,aov:aov3,vsCur:badgeHtml(prevWeek.revenue,cur.revenue)}}
    ];
    let mh="";rows.forEach(r=>{{mh+=`<tr><td>${{r.label}}</td><td class="right">${{fmtFull(Math.round(r.rev))}}</td><td class="right">${{fmtFull(r.ord)}}</td><td class="right">${{fmtFull(r.aov)}}</td><td>${{r.vsCur}}</td></tr>`;}});
    document.getElementById("lookup-mom-table").innerHTML=mh;

    /* (4) Category */
    if(charts["lookup-category-chart"])charts["lookup-category-chart"].destroy();
    charts["lookup-category-chart"]=new Chart(document.getElementById("lookup-category-chart"),{{type:"doughnut",data:{{labels:Object.keys(cur.categories).map(x=>categoryNames[x]+" "+categoryNamesVi[x]),datasets:[{{data:Object.values(cur.categories),backgroundColor:chartColors}}]}},options:{{responsive:true,maintainAspectRatio:false}}}});
    const cats=["san","son","congcu","decor","other"];
    const sortedCats=[...cats].sort((a,b)=>cur.categories[b]-cur.categories[a]);
    let ch2="";sortedCats.forEach(c=>{{
        ch2+=`<tr><td>${{categoryNames[c]}} ${{categoryNamesVi[c]}}</td><td class="right">${{fmtFull(cur.categories[c])}}</td><td class="right">${{fmtFull(prev.categories[c])}}</td><td>${{badgeHtml(cur.categories[c],prev.categories[c])}}</td></tr>`;
    }});
    document.getElementById("lookup-category-table").innerHTML=ch2;

    /* (5) Fees */
    const feesArr=["shopee","tiktok","web","lazada"].map(p=>cur.feesByChannel[p]||0);
    if(charts["lookup-fees-chart"])charts["lookup-fees-chart"].destroy();
    charts["lookup-fees-chart"]=new Chart(document.getElementById("lookup-fees-chart"),{{type:"bar",data:{{labels:chLabels,datasets:[{{label:"平台費用",data:feesArr,backgroundColor:chartColors}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
    let fh="",totalRev=0,totalFees=0,totalNet=0;
    ["shopee","tiktok","web","lazada"].forEach(p=>{{
        const r=cur.byChannel[p]||0,f=cur.feesByChannel[p]||0,nv=cur.netByChannel[p]||0;
        const pct=r?((f/r)*100).toFixed(1):"-";
        fh+=`<tr><td>${{platformNames[p]}}</td><td class="right">${{fmtFull(r)}}</td><td class="right">${{fmtFull(f)}}</td><td class="right">${{pct}}%</td><td class="right">${{fmtFull(nv)}}</td></tr>`;
        totalRev+=r;totalFees+=f;totalNet+=nv;
    }});
    const totalPct=totalRev?((totalFees/totalRev)*100).toFixed(1):"-";
    fh+=`<tr style="font-weight:700;border-top:2px solid #333"><td>合計 Tổng Cộng</td><td class="right">${{fmtFull(totalRev)}}</td><td class="right">${{fmtFull(totalFees)}}</td><td class="right">${{totalPct}}%</td><td class="right">${{fmtFull(totalNet)}}</td></tr>`;
    document.getElementById("lookup-fees-table").innerHTML=fh;

    /* (6) MoM channel - 3 datasets */
    const chPrev=["shopee","tiktok","web","lazada"].map(p=>prev.byChannel[p]||0);
    const chWeek=["shopee","tiktok","web","lazada"].map(p=>prevWeek.byChannel[p]||0);
    if(charts["lookup-mom-channel-chart"])charts["lookup-mom-channel-chart"].destroy();
    charts["lookup-mom-channel-chart"]=new Chart(document.getElementById("lookup-mom-channel-chart"),{{type:"bar",data:{{labels:chLabels,datasets:[{{label:"週前",data:chWeek,backgroundColor:"#9b59b6"}},{{label:"昨日",data:chPrev,backgroundColor:"#95a5a6"}},{{label:"本期",data:chCur,backgroundColor:"#e74c3c"}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
    let mch="";
    ["shopee","tiktok","web","lazada"].forEach(p=>{{
        const a=cur.byChannel[p]||0,b=prev.byChannel[p]||0,c=prevWeek.byChannel[p]||0;
        mch+=`<tr><td>${{platformNames[p]}}</td><td class="right">${{fmtFull(a)}}</td><td class="right">${{fmtFull(b)}}</td><td class="right">${{fmtFull(c)}}</td><td>${{badgeHtml(a,b)}}</td></tr>`;
    }});
    document.getElementById("lookup-mom-channel-table").innerHTML=mch;

    /* (7) Products */
    let ph="";
    cur.products.forEach((p,i)=>{{
        ph+=`<tr><td>${{i+1}}</td><td>${{p.name}}</td><td class="right">${{fmtFull(p.qty)}}</td><td class="right">${{fmtFull(p.revenue)}}</td></tr>`;
    }});
    if(!ph)ph='<tr><td colspan="4" style="text-align:center;color:#999;padding:20px;">Không có dữ liệu sản phẩm trong khoảng đã chọn</td></tr>';
    document.getElementById("lookup-products-table").innerHTML=ph;
    // SP detail co bo dieu khien rieng, khong re-render khi controls main thay doi
}}

/* ===== SP SEARCH ===== */
let selectedSP=null;
let spLookupMode="single";
let spLookupChannel="all";

function getSpLookupRange(){{
    if(spLookupMode==="single"){{
        const v=document.getElementById("sp-date").value||lookupLastDate;
        return [v,v];
    }} else {{
        let s=document.getElementById("sp-start").value||lookupFirstDate;
        let e=document.getElementById("sp-end").value||lookupLastDate;
        if(s>e){{const t=s;s=e;e=t;}}
        return [s,e];
    }}
}}
function getSpLookupPlatforms(){{
    return spLookupChannel==="all"?["shopee","tiktok","web","lazada"]:[spLookupChannel];
}}
function getAllSPNames(){{ return Object.keys(PRODUCTS).sort(); }}

function aggregateSP(spName,start,end,platforms){{
    const data=PRODUCTS[spName];
    if(!data)return null;
    const platformSet=new Set(platforms||["shopee","tiktok","web","lazada"]);
    const result={{name:spName,total_qty:0,total_revenue:0,
        by_channel:{{shopee:{{qty:0,revenue:0}},tiktok:{{qty:0,revenue:0}},web:{{qty:0,revenue:0}},lazada:{{qty:0,revenue:0}}}},
        by_date:{{}},dates_with_sales:0,platforms_filter:Array.from(platformSet)}};
    Object.keys(data).forEach(d=>{{
        if(d>=start&&d<=end){{
            const dayData=data[d];
            let hasData=false;
            if(!result.by_date[d])result.by_date[d]={{qty:0,revenue:0,channels:{{}}}};
            Object.keys(dayData).forEach(p=>{{
                if(!platformSet.has(p))return;
                const v=dayData[p];
                result.total_qty+=v.qty;
                result.total_revenue+=v.revenue;
                if(result.by_channel[p]){{
                    result.by_channel[p].qty+=v.qty;
                    result.by_channel[p].revenue+=v.revenue;
                }}
                result.by_date[d].qty+=v.qty;
                result.by_date[d].revenue+=v.revenue;
                result.by_date[d].channels[p]=v;
                hasData=true;
            }});
            if(hasData)result.dates_with_sales++;
        }}
    }});
    // Bo ngay khong co data
    Object.keys(result.by_date).forEach(d=>{{ if(result.by_date[d].qty===0&&result.by_date[d].revenue===0)delete result.by_date[d]; }});
    return result;
}}

function renderSPDetail(spName){{
    selectedSP=spName;
    const [s,e]=getSpLookupRange();
    const platforms=getSpLookupPlatforms();
    const data=aggregateSP(spName,s,e,platforms);
    if(!data){{
        document.getElementById("lookup-sp-detail").style.display="none";
        return;
    }}
    document.getElementById("lookup-sp-detail").style.display="block";
    const same=(s===e);
    const periodInfo=same?`Ngày <b>${{fmtDate(s)}}</b>`:`Từ <b>${{fmtDate(s)}}</b> đến <b>${{fmtDate(e)}}</b>`;
    const chInfo=spLookupChannel==="all"?"tất cả 4 kênh":platformNames[spLookupChannel];
    document.getElementById("sp-compare-info").innerHTML=`Kỳ: ${{periodInfo}} | Kênh: <b>${{chInfo}}</b>`;
    document.getElementById("lookup-sp-name").innerHTML=`<strong>${{spName}}</strong>`;

    const kpi=[
        {{l:"Tổng SL",v:fmtFull(data.total_qty)}},
        {{l:"Tổng Doanh Thu",v:fmt(data.total_revenue)}},
        {{l:"Số ngày có đơn",v:data.dates_with_sales+" ngày"}},
        {{l:"DT/đơn TB",v:data.total_qty>0?fmt(Math.round(data.total_revenue/data.total_qty)):"-"}},
    ];
    let kpiHtml="";
    kpi.forEach(x=>{{ kpiHtml+=`<div class="kpi-card"><div class="kpi-label">${{x.l}}</div><div class="kpi-value">${{x.v}}</div></div>`; }});
    document.getElementById("lookup-sp-kpi").innerHTML=kpiHtml;

    let chHtml="";
    const channelsToShow=spLookupChannel==="all"?["shopee","tiktok","web","lazada"]:[spLookupChannel];
    channelsToShow.forEach(p=>{{
        const v=data.by_channel[p];
        const pct=data.total_revenue>0?(v.revenue/data.total_revenue*100).toFixed(1):"0.0";
        chHtml+=`<tr><td>${{platformNames[p]}}</td><td class="right">${{fmtFull(v.qty)}}</td><td class="right">${{fmtFull(v.revenue)}}</td><td class="right">${{pct}}%</td></tr>`;
    }});
    if(spLookupChannel==="all"){{
        chHtml+=`<tr style="font-weight:700;border-top:2px solid #333"><td>Tổng</td><td class="right">${{fmtFull(data.total_qty)}}</td><td class="right">${{fmtFull(data.total_revenue)}}</td><td class="right">100%</td></tr>`;
    }}
    document.getElementById("lookup-sp-channel-table").innerHTML=chHtml;

    const sortedDates=Object.keys(data.by_date).sort();
    let dateHtml="";
    sortedDates.forEach(d=>{{
        const v=data.by_date[d];
        const entries=Object.entries(v.channels);
        const main=entries.length>0?entries.sort((a,b)=>b[1].revenue-a[1].revenue)[0]:null;
        const chName=main?platformNames[main[0]]:"—";
        dateHtml+=`<tr><td>${{fmtDate(d)}}</td><td class="right">${{fmtFull(v.qty)}}</td><td class="right">${{fmtFull(v.revenue)}}</td><td>${{chName}}</td></tr>`;
    }});
    if(!dateHtml)dateHtml='<tr><td colspan="4" style="text-align:center;color:#999;padding:20px;">Không có data trong khoảng đã chọn</td></tr>';
    document.getElementById("lookup-sp-date-table").innerHTML=dateHtml;
}}

function handleSPSearch(q){{
    const resultDiv=document.getElementById("lookup-sp-results");
    const ql=q.toLowerCase().trim();
    if(!ql){{ resultDiv.classList.remove("show"); resultDiv.innerHTML=""; return; }}
    const matched=getAllSPNames().filter(n=>n.toLowerCase().includes(ql));
    if(matched.length===0){{
        resultDiv.classList.add("show");
        resultDiv.innerHTML='<div class="search-result-item" style="color:#999">Không tìm thấy SP nào khớp</div>';
        return;
    }}
    let h="";
    matched.slice(0,15).forEach(n=>{{
        const lower=n.toLowerCase();
        const idx=lower.indexOf(ql);
        let display=n;
        if(idx>=0){{ display=n.substring(0,idx)+'<mark>'+n.substring(idx,idx+q.length)+'</mark>'+n.substring(idx+q.length); }}
        const sumDates=Object.keys(PRODUCTS[n]).length;
        h+=`<div class="search-result-item" data-sp="${{n.replace(/"/g,'&quot;')}}">${{display}}<div class="meta">${{sumDates}} ngày có đơn (tổng)</div></div>`;
    }});
    if(matched.length>15){{ h+=`<div class="search-result-item" style="color:#999;text-align:center;cursor:default;">... còn ${{matched.length-15}} SP khác (gõ thêm để lọc)</div>`; }}
    resultDiv.innerHTML=h;
    resultDiv.classList.add("show");
    resultDiv.querySelectorAll(".search-result-item[data-sp]").forEach(el=>{{
        el.addEventListener("click",()=>{{
            const sp=el.dataset.sp;
            document.getElementById("lookup-sp-search").value=sp;
            resultDiv.classList.remove("show");
            renderSPDetail(sp);
        }});
    }});
}}

/* ===== RENDER FUNCTIONS ===== */
function overview(){{
    const all=["shopee","tiktok","web","lazada"];
    renderKpi(all,"tongQuan-kpi");
    renderDaily(all,"tongQuan-daily-chart");
    renderTrend(all,"tongQuan-trend-chart");
    renderChannelCompare("tongQuan-channel-chart");
    renderMomTable(all,"tongQuan-mom-table");
    renderCategoryChart(all,"tongQuan-category-chart");
    renderCategoryTable(all,"tongQuan-category-table");
    renderFeesChart("tongQuan-fees-chart");
    renderFeesTable("tongQuan-fees-table");
    renderMomChannelChart("tongQuan-mom-channel-chart");
    renderMomChannelTable("tongQuan-mom-channel-table");
    renderProducts(all,"tongQuan-products-table",true);
}}

function platform(p){{
    renderKpi([p],p+"-kpi");
    renderDaily([p],p+"-daily-chart");
    renderTrend([p],p+"-trend-chart");
    renderMomTable([p],p+"-mom-table");
    renderCategoryChart([p],p+"-category-chart");
    renderCategoryTable([p],p+"-category-table");
    renderPlatformFees(p,p+"-fees-chart",p+"-fees-table");
    renderOrdersTrend(p,p+"-orders-chart");
    renderProducts([p],p+"-products-table",false);
}}

document.querySelectorAll(".tab-btn").forEach(b=>{{b.addEventListener("click",e=>{{
    document.querySelectorAll(".tab-content").forEach(t=>t.classList.remove("active"));
    const tid=e.target.dataset.tab;
    document.getElementById(tid).classList.add("active");
    document.querySelectorAll(".tab-btn").forEach(t=>t.classList.remove("active"));
    e.target.classList.add("active");
    if(tid==="tong-quan")overview();
    else if(tid==="lookup")lookup();
    else platform(resolveKey(tid));
}});}});

/* Lookup tab event handlers - controls dau tab (data-target="main") */
document.querySelectorAll('#lookup .mode-btn[data-target="main"]').forEach(b=>{{b.addEventListener("click",e=>{{
    lookupMode=e.target.dataset.mode;
    document.querySelectorAll('#lookup .mode-btn[data-target="main"]').forEach(t=>t.classList.remove("active"));
    e.target.classList.add("active");
    document.getElementById("lookup-single-controls").style.display = lookupMode==="single"?"":"none";
    document.getElementById("lookup-range-controls").style.display = lookupMode==="range"?"":"none";
    lookup();
}});}});
document.querySelectorAll('#lookup .ch-pill[data-target="main"]').forEach(b=>{{b.addEventListener("click",e=>{{
    lookupChannel=e.target.dataset.channel;
    document.querySelectorAll('#lookup .ch-pill[data-target="main"]').forEach(t=>t.classList.remove("active"));
    e.target.classList.add("active");
    lookup();
}});}});
["lookup-date","lookup-start","lookup-end"].forEach(id=>{{
    const el=document.getElementById(id);
    if(el)el.addEventListener("change",lookup);
}});

/* SP detail controls event handlers (data-target="sp") */
document.querySelectorAll('#lookup .mode-btn[data-target="sp"]').forEach(b=>{{b.addEventListener("click",e=>{{
    spLookupMode=e.target.dataset.mode;
    document.querySelectorAll('#lookup .mode-btn[data-target="sp"]').forEach(t=>t.classList.remove("active"));
    e.target.classList.add("active");
    document.getElementById("sp-single-controls").style.display = spLookupMode==="single"?"":"none";
    document.getElementById("sp-range-controls").style.display = spLookupMode==="range"?"":"none";
    if(selectedSP)renderSPDetail(selectedSP);
}});}});
document.querySelectorAll('#lookup .ch-pill[data-target="sp"]').forEach(b=>{{b.addEventListener("click",e=>{{
    spLookupChannel=e.target.dataset.channel;
    document.querySelectorAll('#lookup .ch-pill[data-target="sp"]').forEach(t=>t.classList.remove("active"));
    e.target.classList.add("active");
    if(selectedSP)renderSPDetail(selectedSP);
}});}});
["sp-date","sp-start","sp-end"].forEach(id=>{{
    const el=document.getElementById(id);
    if(el)el.addEventListener("change",()=>{{ if(selectedSP)renderSPDetail(selectedSP); }});
}});

/* SP search event handlers */
const spSearchInput=document.getElementById("lookup-sp-search");
if(spSearchInput){{
    spSearchInput.addEventListener("input",e=>handleSPSearch(e.target.value));
    spSearchInput.addEventListener("focus",e=>{{ if(e.target.value)handleSPSearch(e.target.value); }});
}}
document.addEventListener("click",e=>{{
    if(!e.target.closest(".search-bar")){{
        const r=document.getElementById("lookup-sp-results");
        if(r)r.classList.remove("show");
    }}
}});

document.querySelectorAll(".month-btn").forEach(b=>{{b.addEventListener("click",e=>{{
    currentMonth=e.target.dataset.month;
    document.querySelectorAll(".month-btn").forEach(t=>t.classList.remove("active"));
    e.target.classList.add("active");
    const at=document.querySelector(".tab-content.active");
    if(at.id==="tong-quan")overview();else platform(resolveKey(at.id));
}});}});

overview();
    </script>
</body>
</html>'''

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)


def _generate_platform_tabs():
    """Generate HTML for the 4 platform tabs."""
    tabs = ""
    for pid, pname in [("shopee","Shopee"),("tiktok","TikTok"),("lazada","Lazada"),("website","Website")]:
        p = pid if pid != "website" else "web"
        tabs += f'''
        <div id="{pid}" class="tab-content">
            <div class="section-title">① 關鍵績效指標 KPI Tổng Quan</div>
            <div class="kpi-grid" id="{p}-kpi"></div>

            <div class="section-title">② 每日營收圖表 Biểu Đồ Doanh Thu Theo Ngày</div>
            <div class="chart-container"><div class="chart-title">{pname} 每日營收 Doanh thu hàng ngày</div><div class="chart-wrapper"><canvas id="{p}-daily-chart"></canvas></div></div>

            <div class="section-title">③ 各月份營收成長比較 So Sánh Tăng Trưởng Các Tháng</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">{pname} 營收趨勢 Xu hướng doanh thu</div><div class="chart-wrapper"><canvas id="{p}-trend-chart"></canvas></div></div>
                <div class="chart-container"><div class="chart-title">{pname} 訂單趨勢 Xu hướng đơn hàng</div><div class="chart-wrapper"><canvas id="{p}-orders-chart"></canvas></div></div>
            </div>
            <div class="table-container">
                <div class="chart-title">{pname} 各月份營收成長明細 Chi tiết tăng trưởng từng tháng</div>
                <table><thead><tr><th>月份 Tháng</th><th class="right">營收 Doanh Thu</th><th class="right">訂單數 Đơn Hàng</th><th class="right">客單價 AOV</th><th>環比成長 Tăng Trưởng</th></tr></thead><tbody id="{p}-mom-table"></tbody></table>
            </div>

            <div class="section-title">④ 產品類別營收 Doanh Thu Theo Danh Mục</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">{pname} 類別營收分佈 Phân bổ danh mục</div><div class="chart-wrapper"><canvas id="{p}-category-chart"></canvas></div></div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">{pname} 類別營收 Danh mục (本月 vs 上月 Tháng này vs Tháng trước)</div>
                    <table><thead><tr><th>類別 Danh Mục</th><th class="right">本月營收 Tháng Này</th><th class="right">上月營收 Tháng Trước</th><th>變化 Thay Đổi</th></tr></thead><tbody id="{p}-category-table"></tbody></table>
                </div>
            </div>

            <div class="section-title">⑤ 平台費用分析 Chi Phí Sàn &amp; Tỷ Lệ Phí</div>
            <div class="section-grid">
                <div class="chart-container"><div class="chart-title">{pname} 平台費用趨勢 Xu hướng phí sàn</div><div class="chart-wrapper"><canvas id="{p}-fees-chart"></canvas></div></div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">{pname} 費用及費率明細 Chi tiết phí và tỷ lệ</div>
                    <table><thead><tr><th>月份 Tháng</th><th class="right">營收 Doanh Thu</th><th class="right">費用 Phí</th><th class="right">費率 Tỷ Lệ %</th><th class="right">淨收入 DT Ròng</th></tr></thead><tbody id="{p}-fees-table"></tbody></table>
                </div>
            </div>

            <div class="section-title">⑦ 暢銷產品 SP Bán Chạy Theo Doanh Thu</div>
            <div class="table-container">
                <div class="chart-title">{pname} Top 5 暢銷產品 Sản phẩm bán chạy</div>
                <table><thead><tr><th>#</th><th>產品名稱 Tên SP</th><th class="right">數量 SL</th><th class="right">營收 Doanh Thu</th></tr></thead><tbody id="{p}-products-table"></tbody></table>
            </div>
        </div>
'''
    return tabs


def main():
    output_path = sys.argv[1] if len(sys.argv) > 1 else "/sessions/nice-funny-mccarthy/mnt/Tồn- xuất kho- hàng hóa/Dashboard_Loho_House_2026.html"

    print("=== Loho House Dashboard Updater v2 ===")
    print(f"Output: {output_path}")
    print()

    print("📥 Downloading 8 sheets...")
    csv_paths = {}
    for name, info in SHEETS.items():
        print(f"  → {name} (gid={info['gid']})...", end=" ")
        path = download_csv(name, info["gid"])
        if path:
            csv_paths[name] = path
            lines = sum(1 for _ in open(path, "r", encoding="utf-8", errors="replace"))
            print(f"OK ({lines} lines)")
        else:
            print("FAILED")

    print("\n📊 Processing revenue data...")
    platforms_data = {}
    daily_data = {}
    for platform in ["shopee", "tiktok", "web", "lazada"]:
        if platform not in csv_paths:
            platforms_data[platform] = {}
            daily_data[platform] = {}
            continue
        m_data, d_data = process_platform(platform, csv_paths[platform])
        platforms_data[platform] = m_data
        daily_data[platform] = d_data
        for mk in sorted(m_data.keys(), key=lambda x: int(x[1:])):
            m = m_data[mk]
            print(f"  {platform} {mk}: {m['orders']} orders, revenue={int(m['revenue']):,}")
        print(f"  {platform} daily: {len(d_data)} dates")

    print("\n🏷️  Processing categories...")
    categories_data = {}
    daily_categories_data = {}
    for platform in ["shopee", "tiktok", "web", "lazada"]:
        raw_name = platform + "_raw"
        if raw_name not in csv_paths:
            categories_data[platform] = {}
            daily_categories_data[platform] = {}
            continue
        monthly_rev = {k: v["revenue"] for k, v in platforms_data[platform].items()}
        cats, cats_daily = process_raw_for_categories(platform, csv_paths[raw_name], monthly_rev)
        categories_data[platform] = cats
        daily_categories_data[platform] = cats_daily
        for mk in sorted(cats.keys(), key=lambda x: int(x[1:])):
            c = cats[mk]
            print(f"  {platform} {mk}: san={c.get('san',0):,} son={c.get('son',0):,} congcu={c.get('congcu',0):,} decor={c.get('decor',0):,}")

    # Enrich products from raw sheets for platforms with empty product names
    print("\n📦 Enriching products from raw sheets...")
    for platform in ["web", "lazada"]:
        raw_name = platform + "_raw"
        if raw_name in csv_paths:
            enrich_products_from_raw(platform, csv_paths[raw_name], platforms_data, daily_data)

    print("\n🔧 Building data JSON...")
    data_json = build_data_json(platforms_data, categories_data)
    daily_json = build_daily_json(daily_data, daily_categories_data)
    products_json = build_products_index(daily_data)
    print(f"  Built products index: {len(products_json)} unique products")

    print(f"\n📝 Generating HTML → {output_path}")
    generate_html(data_json, daily_json, products_json, output_path)

    file_size = os.path.getsize(output_path)
    print(f"\n✅ Done! File size: {file_size:,} bytes")
    print(f"📅 Updated: {datetime.now().strftime('%d/%m/%Y %H:%M')}")


if __name__ == "__main__":
    main()
