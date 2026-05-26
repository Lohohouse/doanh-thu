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
try:
    from zoneinfo import ZoneInfo
    VN_TZ = ZoneInfo('Asia/Ho_Chi_Minh')
except ImportError:
    VN_TZ = None

# ===== CONFIG =====
FILE_ID = "1jwPEzRMcoYBJywZkW4Vn8dKe_w5hU-zVq51zSoXGY0M"
BASE_URL = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=csv&gid="

# Google Sheet để nhập tay data báo cáo tuần
WEEKLY_SHEET_ID = "1coZ2UmG8blAfgAwR5Ya8wg2l1BD9DJeHfu9yQCsT4Oo"
WEEKLY_SHEET_GID = "0"
WEEKLY_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{WEEKLY_SHEET_ID}/export?format=csv&gid={WEEKLY_SHEET_GID}"

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


def _norm_label(s):
    """Normalize Vietnamese label for matching: lowercase, strip, collapse spaces, remove diacritics."""
    if not s: return ""
    import unicodedata
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.lower().strip()
    # Vietnamese đ / Đ are separate characters — NFD does not decompose them
    s = s.replace("đ", "d").replace("Đ", "d")
    # Remove punctuation that may vary
    s = s.replace(":", "").replace("(", "").replace(")", "").replace("%", "").replace(",", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _parse_weekly_csv(csv_text):
    """Parse the weekly report Google Sheet CSV into the JSON structure used by the dashboard.

    Sheet uses friendly Vietnamese labels in column A. Each section has its own column meaning:
        ① Doanh thu quyết toán: B=Shopee, C=Tiktok, D=Lazada, E=Website
        ② QC tuần: B=Shopee tuần trước, C=Shopee tuần này, D=Tiktok tuần trước, E=Tiktok tuần này
        ③ QC tháng Shopee: B=Tháng này, C=Cùng kỳ, D=Tháng sau
        ④ QC tháng Tiktok: B=Tháng này, C=Cùng kỳ, D=Tháng sau
        ⑤ Chăm sóc KH: B=Shopee, C=Tiktok, D=Lazada, E=Zalo OA
        ⑥ Đánh giá: B=Shopee, C=Tiktok, D=Lazada
        ⑦ Khiếu nại: B=Shopee, C=Tiktok, D=Lazada, E=Website
        ⑧ Tỷ lệ KN: B=Shopee, C=Tiktok, D=Lazada, E=Website

    Parser tracks current section based on header rows containing ①/②/③/...
    """
    def num(v):
        if v is None: return 0
        v = str(v).strip()
        if not v or v in ("—", "-", "/", "N/A"): return 0
        v = v.replace(",", "").replace(" ", "").replace("đ", "").replace("%", "")
        try: return float(v)
        except: return 0

    def num_or_none(v):
        if v is None: return None
        v = str(v).strip()
        if not v or v in ("—", "-", "/", "N/A"): return None
        v = v.replace(",", "").replace(" ", "").replace("đ", "").replace("%", "")
        try: return float(v)
        except: return None

    def s(v):
        return (str(v) if v is not None else "").strip()

    # Parse all rows
    raw_rows = list(csv.reader(io.StringIO(csv_text)))

    # Section tracker
    section = ""
    rows_by_section = {}  # section -> list of (label_norm, original_label, [B,C,D,E])

    for r in raw_rows:
        if not r or all(not c.strip() for c in r):
            continue
        col_a = (r[0] or "").strip()
        # Pad to 6 cols
        r = (r + [""] * 6)[:6]
        cols_bcde = [r[1], r[2], r[3], r[4]]

        # Detect section header
        if any(col_a.startswith(prefix) for prefix in ["①","②","③","④","⑤","⑥","⑦","⑧","⑨"]):
            # Map number to section key
            section_map = {
                "①": "settlement", "②": "ads_weekly",
                "③": "ads_monthly_shopee", "④": "ads_monthly_tiktok",
                "⑤": "customer_care", "⑥": "reviews",
                "⑦": "complaints", "⑧": "complaint_ratio",
                "⑨": "extra",
            }
            section = section_map.get(col_a[0], "")
            rows_by_section.setdefault(section, [])
            continue
        # Detect THONG TIN TUAN section
        if "THONG TIN TUAN" in _norm_label(col_a).upper() or "thong tin tuan" in _norm_label(col_a):
            section = "meta"
            rows_by_section.setdefault(section, [])
            continue
        # Skip title/instruction rows
        if not col_a or col_a.startswith("📋") or col_a.startswith("📅") or col_a.startswith("📝") or col_a.startswith("BÁO CÁO"):
            continue
        # Skip column-header subhead rows (when label is generic)
        norm_label = _norm_label(col_a)
        if norm_label in ("chi tieu", "muc", "key", "name", "label"):
            continue

        if section:
            rows_by_section.setdefault(section, []).append((norm_label, col_a, cols_bcde))

    def get_row(sec, *label_patterns):
        """Find row in section whose normalized label matches any pattern (substring match)."""
        for norm_label, orig_label, cols in rows_by_section.get(sec, []):
            for p in label_patterns:
                if _norm_label(p) in norm_label:
                    return cols
        return ["", "", "", ""]

    # === META ===
    label = s(get_row("meta", "ten tuan")[0]) or "Tuần hiện tại"
    code = s(get_row("meta", "ma tuan")[0]) or "2026-W00"

    # === SETTLEMENT ===
    set_row = get_row("settlement", "doanh thu quyet toan", "quyet toan")
    settlement = {
        "shopee_settled":  num(set_row[0]),
        "tiktok_settled":  num(set_row[1]),
        "lazada_settled":  num(set_row[2]),
        "website_settled": num(set_row[3]),
    }

    # === ADS WEEKLY === B=Shopee prev, C=Shopee cur, D=Tiktok prev, E=Tiktok cur
    cost_r = get_row("ads_weekly", "chi phi quang cao", "chi phi qc")
    rev_r  = get_row("ads_weekly", "doanh so tu quang cao", "doanh so qc", "doanh so tu qc")
    ord_r  = get_row("ads_weekly", "so don hang tu qc", "so don hang", "don hang tu qc")
    sp_r   = get_row("ads_weekly", "so san pham ban tu qc", "san pham ban", "so sp ban")
    ads_weekly = {
        "shopee": {
            "prev_cost":     num_or_none(cost_r[0]),
            "cost":          num_or_none(cost_r[1]),
            "prev_revenue":  num_or_none(rev_r[0]),
            "revenue":       num_or_none(rev_r[1]),
            "prev_orders":   num_or_none(ord_r[0]),
            "orders":        num_or_none(ord_r[1]),
            "prev_products": num_or_none(sp_r[0]),
            "products":      num_or_none(sp_r[1]),
        },
        "tiktok": {
            "prev_cost":     num_or_none(cost_r[2]),
            "cost":          num_or_none(cost_r[3]),
            "prev_revenue":  num_or_none(rev_r[2]),
            "revenue":       num_or_none(rev_r[3]),
            "prev_orders":   num_or_none(ord_r[2]),
            "orders":        num_or_none(ord_r[3]),
            "prev_products": num_or_none(sp_r[2]),
            "products":      num_or_none(sp_r[3]),
        },
    }

    # === ADS MONTHLY === Sections 3 (shopee) and 4 (tiktok) — B=Tháng này, C=Cùng kỳ, D=Tháng sau
    def parse_monthly(sec):
        cp = get_row(sec, "chi phi qc thang", "chi phi qc", "chi phi quang cao")
        ds = get_row(sec, "doanh so qc thang", "doanh so qc", "doanh so quang cao")
        dh = get_row(sec, "so don hang thang", "don hang thang", "so don hang")
        sp = get_row(sec, "so sp ban thang", "sp ban thang", "san pham ban")
        td = get_row(sec, "tong doanh thu ban hang", "tong doanh thu")
        return {
            "t_current_label":     "Tháng này",
            "t_prev_year_label":   "Cùng kỳ",
            "t_next_label":        "Tháng sau",
            "t_current_cost":      num_or_none(cp[0]),
            "t_prev_year_cost":    num_or_none(cp[1]),
            "t_next_cost":         num_or_none(cp[2]),
            "t_current_revenue":   num_or_none(ds[0]),
            "t_prev_year_revenue": num_or_none(ds[1]),
            "t_next_revenue":      num_or_none(ds[2]),
            "t_current_orders":    num_or_none(dh[0]),
            "t_prev_year_orders":  num_or_none(dh[1]),
            "t_next_orders":       num_or_none(dh[2]),
            "t_current_products":  num_or_none(sp[0]),
            "t_prev_year_products":num_or_none(sp[1]),
            "t_next_products":     num_or_none(sp[2]),
            "t_current_total_rev": num_or_none(td[0]),
            "t_next_total_rev":    num_or_none(td[2]),
        }
    ads_monthly_sh = parse_monthly("ads_monthly_shopee")
    ads_monthly_tk = parse_monthly("ads_monthly_tiktok")
    ads_monthly = {"shopee": ads_monthly_sh, "tiktok": ads_monthly_tk}

    # === TONG_DT === (Pulled from monthly sections row "Tổng doanh thu bán hàng")
    ads_total_revenue = {
        "shopee_t4": ads_monthly_sh.get("t_current_total_rev") or 0,
        "shopee_t5": ads_monthly_sh.get("t_next_total_rev") or 0,
        "tiktok_t4": ads_monthly_tk.get("t_current_total_rev") or 0,
        "tiktok_t5": ads_monthly_tk.get("t_next_total_rev") or 0,
    }

    # === CUSTOMER CARE === B=Shopee, C=Tiktok, D=Lazada, E=Zalo OA
    chat_r = get_row("customer_care", "luot chat")
    conv_r = get_row("customer_care", "ty le chuyen doi")
    rev_r2 = get_row("customer_care", "doanh so tu chat", "doanh so")
    care_channels = ["Shopee", "Tiktok", "Lazada", "Zalo OA - Zalo shop"]
    customer_care = []
    for i, ch in enumerate(care_channels):
        customer_care.append({
            "channel": ch,
            "chats": num_or_none(chat_r[i]),
            "conversion": num_or_none(conv_r[i]),
            "revenue": num_or_none(rev_r2[i]),
        })

    # === REVIEWS === B=Shopee, C=Tiktok, D=Lazada
    r5 = get_row("reviews", "rate 5 sao")
    r3 = get_row("reviews", "rate 3 sao")
    r1 = get_row("reviews", "rate 1 sao")
    review_channels = ["Shopee", "Tiktok", "Lazada"]
    reviews = []
    for i, ch in enumerate(review_channels):
        # Reason and status are per-channel labeled rows
        reason_row = get_row("reviews", f"ly do rate thap - {ch.lower()}", f"ly do {ch.lower()}")
        status_row = get_row("reviews", f"tinh trang xu ly - {ch.lower()}", f"tinh trang {ch.lower()}")
        reason = s(reason_row[0]).replace(";", "\n")
        status = s(status_row[0])
        reviews.append({
            "channel": ch,
            "rate5": num_or_none(r5[i]),
            "rate3": num_or_none(r3[i]),
            "rate1": num_or_none(r1[i]),
            "reason": reason,
            "status": status,
        })

    # === COMPLAINTS === B=Shopee, C=Tiktok, D=Lazada, E=Website
    tm = get_row("complaints", "tong kn trong thang", "tong khieu nai", "tong thang")
    tw = get_row("complaints", "kn trong tuan", "khieu nai trong tuan", "trong tuan")
    dr = get_row("complaints", "da xu ly")
    de = get_row("complaints", "dang xu ly")
    complaint_channels = ["Shopee", "Tiktok", "Lazada", "Website"]
    complaints = []
    for i, ch in enumerate(complaint_channels):
        reason_row = get_row("complaints", f"ly do khieu nai - {ch.lower()}", f"ly do {ch.lower()}")
        reason = s(reason_row[0]).replace(";", "\n")
        complaints.append({
            "channel": ch,
            "total_month": num_or_none(tm[i]),
            "this_week": num_or_none(tw[i]),
            "resolved": num_or_none(dr[i]),
            "pending": num_or_none(de[i]),
            "reason": reason,
        })

    # === COMPLAINT RATIO === B=Shopee, C=Tiktok, D=Lazada, E=Website
    sk = get_row("complaint_ratio", "so khieu nai")
    td = get_row("complaint_ratio", "tong don hang", "tong don")
    complaint_ratio = {
        "shopee":  {"complaints": num_or_none(sk[0]), "total_orders": num_or_none(td[0])},
        "tiktok":  {"complaints": num_or_none(sk[1]), "total_orders": num_or_none(td[1])},
        "lazada":  {"complaints": num_or_none(sk[2]), "total_orders": num_or_none(td[2])},
        "website": {"complaints": num_or_none(sk[3]), "total_orders": num_or_none(td[3])},
    }

    return {
        "current_week": code,
        "weeks": {
            code: {
                "label": label,
                "settlement": settlement,
                "ads_weekly": ads_weekly,
                "ads_monthly": ads_monthly,
                "ads_total_revenue": ads_total_revenue,
                "customer_care": customer_care,
                "reviews": reviews,
                "complaints": complaints,
                "complaint_ratio": complaint_ratio,
            }
        }
    }


def load_weekly_report_data():
    """Load manual weekly report data — try Google Sheet first, then local JSON fallback."""
    # Try Google Sheet
    try:
        print(f"  → Fetching from Google Sheet (id={WEEKLY_SHEET_ID})...")
        result = subprocess.run(
            ["curl", "-sL", "-A", "Mozilla/5.0", WEEKLY_SHEET_URL],
            capture_output=True, text=True, timeout=30
        )
        csv_text = result.stdout
        if csv_text and len(csv_text) > 50 and ("Doanh thu quyết toán" in csv_text or "quyet toan" in csv_text.lower() or "①" in csv_text):
            data = _parse_weekly_csv(csv_text)
            print(f"  ✅ Loaded from Google Sheet — Tuần: {list(data['weeks'].values())[0]['label']}")
            return data
        else:
            print(f"  ⚠️ Google Sheet trống hoặc thiếu cấu trúc — fallback sang JSON local")
    except Exception as e:
        print(f"  ⚠️ Lỗi fetch Google Sheet: {e} — fallback sang JSON local")

    # Fallback: local JSON
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        "weekly_report_data.json",
        os.path.join(script_dir, "weekly_report_data.json"),
        os.path.join(os.getcwd(), "weekly_report_data.json"),
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    print(f"  → Loaded weekly report data from local: {path}")
                    return json.load(f)
            except Exception as e:
                print(f"  ⚠️ Error loading {path}: {e}")
                continue
    print("  ⚠️ Không tìm thấy data — Tab Báo Cáo Tuần sẽ trống")
    return {"current_week": "", "weeks": {}}


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


def generate_html(data_json, daily_json, products_json, output_path, weekly_data=None):
    today = datetime.now(VN_TZ).strftime("%d/%m/%Y %H:%M") if VN_TZ else datetime.now().strftime("%d/%m/%Y %H:%M")
    months = get_available_months(data_json)
    last_month = months[-1] if months else "T1"
    data_str = json.dumps(data_json, ensure_ascii=False)
    daily_str = json.dumps(daily_json, ensure_ascii=False)
    products_str = json.dumps(products_json, ensure_ascii=False)
    weekly_str = json.dumps(weekly_data or {"current_week": "", "weeks": {}}, ensure_ascii=False)
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
        :root {{
            --bg-page: #F5EFE6;
            --bg-card: #FFFCF7;
            --bg-soft: #FAF6EE;
            --bg-section: #F9F3E8;
            --primary: #8B6F47;
            --primary-dark: #6B5236;
            --primary-light: #B89970;
            --accent: #C9A961;
            --accent-dark: #A6864B;
            --text-dark: #3D2E1F;
            --text-mid: #6B5236;
            --text-soft: #9B8975;
            --border: #E8DFD3;
            --border-soft: #F0E7DA;
            --header-grad-start: #3D2E1F;
            --header-grad-end: #6B5236;
            --shadow-sm: 0 2px 8px rgba(61,46,31,0.06);
            --shadow-md: 0 4px 16px rgba(61,46,31,0.08);
            --shadow-lg: 0 8px 24px rgba(61,46,31,0.10);
            --green-up: #6B8E5A;
            --green-up-bg: #E6EFDD;
            --red-down: #B5573D;
            --red-down-bg: #F5E1D6;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: var(--bg-page); color: var(--text-dark); }}
        .header {{ background: linear-gradient(135deg, var(--header-grad-start) 0%, var(--header-grad-end) 100%); color: #F5EFE6; padding: 36px 20px; text-align: center; box-shadow: var(--shadow-md); position: relative; overflow: hidden; }}
        .header::before {{ content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, transparent, var(--accent), transparent); }}
        .header h1 {{ font-size: 2.4em; margin-bottom: 10px; font-weight: 600; letter-spacing: 0.5px; }}
        .header h1 .accent {{ color: var(--accent); }}
        .header p {{ font-size: 1.05em; opacity: 0.85; letter-spacing: 0.3px; }}
        .controls {{ background: var(--bg-card); padding: 22px 20px; display: flex; justify-content: center; gap: 10px; flex-wrap: wrap; box-shadow: var(--shadow-sm); border-bottom: 1px solid var(--border); }}
        .month-btn {{ padding: 10px 22px; border: 1.5px solid var(--border); background: var(--bg-card); color: var(--text-mid); border-radius: 24px; cursor: pointer; font-weight: 600; transition: all 0.25s ease; letter-spacing: 0.3px; }}
        .month-btn:hover {{ border-color: var(--primary-light); color: var(--primary); background: var(--bg-soft); }}
        .month-btn.active {{ background: var(--primary); color: var(--bg-card); border-color: var(--primary); box-shadow: 0 2px 8px rgba(139,111,71,0.25); }}
        .tabs {{ display: flex; background: var(--bg-card); border-bottom: 1px solid var(--border); padding: 0 20px; gap: 0; box-shadow: var(--shadow-sm); }}
        .tab-btn {{ padding: 16px 26px; background: none; border: none; cursor: pointer; font-weight: 600; color: var(--text-soft); border-bottom: 3px solid transparent; transition: all 0.25s ease; letter-spacing: 0.3px; }}
        .tab-btn:hover {{ color: var(--primary); background: var(--bg-soft); }}
        .tab-btn.active {{ color: var(--primary-dark); border-bottom-color: var(--accent); background: var(--bg-soft); }}
        .container {{ max-width: 1400px; margin: 24px auto; padding: 0 20px; }}
        .tab-content {{ display: none; }}
        .tab-content.active {{ display: block; }}
        .section-title {{ font-size: 1.25em; font-weight: 700; color: var(--text-dark); margin: 28px 0 16px 0; padding: 0 0 10px 0; border-bottom: 2px solid var(--accent); display: inline-block; letter-spacing: 0.3px; }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 18px; margin-bottom: 28px; }}
        .kpi-card {{ background: var(--bg-card); padding: 22px; border-radius: 12px; box-shadow: var(--shadow-sm); border: 1px solid var(--border-soft); transition: all 0.25s ease; }}
        .kpi-card:hover {{ transform: translateY(-2px); box-shadow: var(--shadow-md); }}
        .kpi-label {{ color: var(--text-soft); font-size: 0.82em; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 0.8px; font-weight: 600; }}
        .kpi-value {{ font-size: 1.85em; font-weight: 700; color: var(--text-dark); margin-bottom: 8px; letter-spacing: -0.3px; }}
        .kpi-change {{ font-size: 0.85em; font-weight: 600; display: flex; align-items: center; gap: 5px; }}
        .chart-container {{ background: var(--bg-card); padding: 22px; border-radius: 12px; box-shadow: var(--shadow-sm); border: 1px solid var(--border-soft); margin-bottom: 20px; position: relative; height: 410px; transition: all 0.25s ease; }}
        .chart-container:hover {{ box-shadow: var(--shadow-md); }}
        .chart-title {{ font-size: 1.05em; font-weight: 700; margin-bottom: 16px; color: var(--text-dark); letter-spacing: 0.2px; }}
        .chart-wrapper {{ position: relative; height: 350px; }}
        .section-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }}
        .table-container {{ background: var(--bg-card); padding: 22px; border-radius: 12px; box-shadow: var(--shadow-sm); border: 1px solid var(--border-soft); overflow-x: auto; margin-bottom: 20px; transition: all 0.25s ease; }}
        .table-container:hover {{ box-shadow: var(--shadow-md); }}
        table {{ width: 100%; border-collapse: collapse; }}
        thead {{ background-color: var(--bg-section); border-bottom: 2px solid var(--accent); }}
        th {{ padding: 13px 14px; text-align: left; font-weight: 700; color: var(--text-dark); font-size: 0.88em; letter-spacing: 0.5px; text-transform: uppercase; }}
        td {{ padding: 13px 14px; border-bottom: 1px solid var(--border-soft); color: var(--text-dark); }}
        tbody tr:hover {{ background-color: var(--bg-soft); }}
        tbody tr:last-child td {{ border-bottom: none; }}
        .badge {{ display: inline-block; padding: 4px 11px; border-radius: 12px; font-size: 0.82em; font-weight: 700; letter-spacing: 0.3px; }}
        .badge.up {{ background-color: var(--green-up-bg); color: var(--green-up); }}
        .badge.down {{ background-color: var(--red-down-bg); color: var(--red-down); }}
        .badge.neutral {{ background-color: var(--border); color: var(--text-soft); }}
        th.right, td.right {{ text-align: right; font-variant-numeric: tabular-nums; }}
        @media (max-width: 768px) {{
            .kpi-grid {{ grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); }}
            .section-grid {{ grid-template-columns: 1fr; }}
            .header h1 {{ font-size: 1.7em; }}
            .tabs {{ overflow-x: auto; }}
        }}
        .lookup-bar {{ background: var(--bg-card); padding: 20px 22px; border-radius: 12px; box-shadow: var(--shadow-sm); border: 1px solid var(--border-soft); margin-bottom: 20px; display: flex; gap: 18px; align-items: center; flex-wrap: wrap; }}
        .lookup-bar label {{ font-weight: 600; color: var(--text-dark); font-size: 0.92em; }}
        .lookup-bar select, .lookup-bar input {{ padding: 9px 13px; border: 1.5px solid var(--border); border-radius: 8px; font-size: 0.93em; font-family: inherit; color: var(--text-dark); background: var(--bg-card); }}
        .lookup-bar input:focus, .lookup-bar select:focus {{ outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(139,111,71,0.12); }}
        .lookup-bar .mode-toggle {{ display: flex; gap: 6px; }}
        .lookup-bar .mode-btn {{ padding: 8px 16px; border: 1.5px solid var(--border); background: var(--bg-card); color: var(--text-mid); border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 0.88em; transition: all 0.2s ease; }}
        .lookup-bar .mode-btn:hover {{ border-color: var(--primary-light); }}
        .lookup-bar .mode-btn.active {{ background: var(--primary); color: var(--bg-card); border-color: var(--primary); }}
        .lookup-bar .channel-pills {{ display: flex; gap: 6px; flex-wrap: wrap; }}
        .lookup-bar .ch-pill {{ padding: 7px 14px; border: 1.5px solid var(--border); background: var(--bg-card); color: var(--text-mid); border-radius: 20px; cursor: pointer; font-size: 0.85em; font-weight: 600; transition: all 0.2s ease; }}
        .lookup-bar .ch-pill:hover {{ border-color: var(--primary-light); color: var(--primary); }}
        .lookup-bar .ch-pill.active {{ background: var(--text-dark); color: var(--bg-card); border-color: var(--text-dark); }}
        .lookup-bar .compare-info {{ font-size: 0.85em; color: var(--text-mid); margin-left: auto; }}
        .search-bar {{ position: relative; margin-bottom: 20px; }}
        .search-bar input {{ width: 100%; padding: 13px 18px; border: 1.5px solid var(--border); border-radius: 10px; font-size: 15px; font-family: inherit; box-sizing: border-box; background: var(--bg-card); color: var(--text-dark); }}
        .search-bar input:focus {{ outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(139,111,71,0.12); }}
        .search-results {{ position: absolute; top: calc(100% + 2px); left: 0; right: 0; background: var(--bg-card); border: 1px solid var(--border); border-radius: 10px; max-height: 320px; overflow-y: auto; z-index: 100; display: none; box-shadow: var(--shadow-lg); }}
        .search-results.show {{ display: block; }}
        .search-result-item {{ padding: 11px 18px; cursor: pointer; border-bottom: 1px solid var(--border-soft); font-size: 14px; transition: background 0.15s; color: var(--text-dark); }}
        .search-result-item:hover {{ background: var(--bg-soft); }}
        .search-result-item:last-child {{ border-bottom: none; }}
        .search-result-item mark {{ background: #F5E2B5; color: var(--primary-dark); padding: 0 3px; border-radius: 3px; font-weight: 600; }}
        .search-result-item .meta {{ font-size: 12px; color: var(--text-soft); margin-top: 3px; }}
        #lookup-sp-detail {{ background: var(--bg-soft); padding: 22px; border-radius: 12px; margin-bottom: 20px; border: 1px solid var(--border-soft); }}
        #lookup-sp-name {{ font-size: 1.05em; color: var(--text-dark); margin-bottom: 18px; padding: 14px; background: var(--bg-card); border-radius: 8px; border-left: 4px solid var(--accent); box-shadow: var(--shadow-sm); }}
        /* === Weekly Report styles === */
        .wr-section {{ background: var(--bg-card); padding: 24px; border-radius: 12px; box-shadow: var(--shadow-sm); border: 1px solid var(--border-soft); margin-bottom: 20px; }}
        .wr-section-title {{ font-size: 1.1em; font-weight: 700; color: var(--text-dark); margin-bottom: 18px; padding-bottom: 10px; border-bottom: 2px solid var(--accent); display: inline-block; letter-spacing: 0.3px; }}
        .wr-meta {{ font-size: 0.85em; color: var(--text-mid); margin-bottom: 14px; font-style: italic; }}
        .wr-grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }}
        .wr-pill {{ display: inline-block; padding: 4px 10px; border-radius: 10px; font-size: 0.8em; font-weight: 600; background: var(--bg-section); color: var(--primary-dark); margin-left: 6px; }}
        .wr-empty {{ text-align: center; padding: 24px; color: var(--text-soft); font-style: italic; background: var(--bg-soft); border-radius: 10px; }}
        @media (max-width: 768px) {{
            .wr-grid-2 {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1><span class="accent">LOHO</span> House 營收報表 Báo Cáo Doanh Thu 2026</h1>
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
        <button class="tab-btn" data-tab="weekly">週報 Báo Cáo Hàng Tuần</button>
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

        <!-- ===== BAO CAO HANG TUAN TAB ===== -->
        <div id="weekly" class="tab-content">
            <div class="lookup-bar">
                <div class="mode-toggle">
                    <button class="mode-btn active" data-mode="single" data-target="wr">1 ngày</button>
                    <button class="mode-btn" data-mode="range" data-target="wr">Khoảng ngày</button>
                </div>
                <div id="wr-single-controls">
                    <label>Chọn ngày:</label>
                    <input type="date" id="wr-date" value="{last_date}" min="{first_date}" max="{last_date}">
                </div>
                <div id="wr-range-controls" style="display:none;">
                    <label>Từ:</label>
                    <input type="date" id="wr-start" value="{first_date}" min="{first_date}" max="{last_date}">
                    <label>Đến:</label>
                    <input type="date" id="wr-end" value="{last_date}" min="{first_date}" max="{last_date}">
                </div>
                <div class="channel-pills">
                    <button class="ch-pill active" data-channel="all" data-target="wr">Tất cả</button>
                    <button class="ch-pill" data-channel="shopee" data-target="wr">Shopee</button>
                    <button class="ch-pill" data-channel="tiktok" data-target="wr">TikTok</button>
                    <button class="ch-pill" data-channel="web" data-target="wr">Web</button>
                    <button class="ch-pill" data-channel="lazada" data-target="wr">Lazada</button>
                </div>
                <div class="compare-info" id="wr-compare-info"></div>
            </div>

            <div class="section-title">① 營收總覽 Doanh Thu (Dashboard thống kê &amp; Quyết toán)</div>
            <div class="table-container">
                <div class="wr-meta" id="wr-revenue-meta"></div>
                <table>
                    <thead>
                        <tr>
                            <th>通路 Kênh</th>
                            <th class="right">營收 Doanh thu từ sàn</th>
                            <th class="right">結算 Doanh thu quyết toán</th>
                            <th class="right">差額 Chênh lệch</th>
                            <th class="right">結算率 % Quyết toán</th>
                        </tr>
                    </thead>
                    <tbody id="wr-revenue-table"></tbody>
                </table>
            </div>

            <div class="section-title">② 暢銷產品 Top 10 SP Bán Chạy (Tuần &amp; Tháng)</div>
            <div class="wr-grid-2">
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">🏆 Top 10 SP trong tuần (kỳ đã chọn) <span class="wr-pill" id="wr-top-week-meta"></span></div>
                    <table>
                        <thead><tr><th>#</th><th>Tên SP</th><th class="right">SL</th><th class="right">Doanh thu</th><th>So với cùng kỳ tuần trước</th></tr></thead>
                        <tbody id="wr-top-week-table"></tbody>
                    </table>
                </div>
                <div class="table-container" style="height:auto;">
                    <div class="chart-title">🏆 Top 10 SP trong tháng <span class="wr-pill" id="wr-top-month-meta"></span></div>
                    <table>
                        <thead><tr><th>#</th><th>Tên SP</th><th class="right">SL</th><th class="right">Doanh thu</th><th>So với tháng trước</th></tr></thead>
                        <tbody id="wr-top-month-table"></tbody>
                    </table>
                </div>
            </div>

            <div class="section-title">③ 廣告費用週報 Chi Phí Quảng Cáo Theo Tuần</div>
            <div class="table-container">
                <div class="wr-meta">Số liệu nhập tay từ file <code>weekly_report_data.json</code> — Tuần đang xem: <b id="wr-ad-week-label">—</b></div>
                <table>
                    <thead>
                        <tr>
                            <th>項目 Kênh bán</th>
                            <th class="right">Shopee tuần trước</th>
                            <th class="right">Shopee tuần này</th>
                            <th>變化 ↕</th>
                            <th class="right">Tiktok tuần trước</th>
                            <th class="right">Tiktok tuần này</th>
                            <th>變化 ↕</th>
                        </tr>
                    </thead>
                    <tbody id="wr-ads-weekly-table"></tbody>
                </table>
            </div>

            <div class="section-title">④ 廣告費用月報 Chi Phí Quảng Cáo Theo Tháng</div>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>項目 Kênh bán</th>
                            <th class="right">Shopee T này</th>
                            <th class="right">Shopee Cùng Kỳ</th>
                            <th class="right">Shopee T sau</th>
                            <th class="right">Tiktok T này</th>
                            <th class="right">Tiktok Cùng Kỳ</th>
                            <th class="right">Tiktok T sau</th>
                        </tr>
                    </thead>
                    <tbody id="wr-ads-monthly-table"></tbody>
                </table>
            </div>

            <div class="section-title">⑤ 廣告費用占比 Tỷ Lệ Phí QC / Doanh Thu</div>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>項目 Chỉ số</th>
                            <th class="right">Shopee T này</th>
                            <th class="right">Shopee T sau</th>
                            <th class="right">Tiktok T này</th>
                            <th class="right">Tiktok T sau</th>
                        </tr>
                    </thead>
                    <tbody id="wr-ads-ratio-table"></tbody>
                </table>
            </div>

            <div class="section-title">⑥ 客戶關懷 Chăm Sóc Khách Hàng (Theo Tuần)</div>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>通路 Kênh</th>
                            <th class="right">Lượt Chat</th>
                            <th class="right">Tỷ lệ chuyển đổi</th>
                            <th class="right">Doanh số</th>
                        </tr>
                    </thead>
                    <tbody id="wr-care-table"></tbody>
                </table>
            </div>

            <div class="section-title">⑦ 客戶評價 Đánh Giá Khách Hàng</div>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>通路 Kênh</th>
                            <th class="right">Rate 5 sao ⭐</th>
                            <th class="right">Rate 3 sao</th>
                            <th class="right">Rate 1 sao</th>
                            <th>Lý do</th>
                            <th>Tình trạng xử lý</th>
                        </tr>
                    </thead>
                    <tbody id="wr-reviews-table"></tbody>
                </table>
            </div>

            <div class="section-title">⑧ 投訴處理 Xử Lý Khiếu Nại</div>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>通路 Kênh</th>
                            <th class="right">Tổng Tháng</th>
                            <th class="right">KN trong tuần</th>
                            <th class="right">Đã xử lý</th>
                            <th class="right">Đang xử lý</th>
                            <th>Lý do</th>
                        </tr>
                    </thead>
                    <tbody id="wr-complaints-table"></tbody>
                </table>
            </div>

            <div class="section-title">⑨ 投訴占比 Tỷ Lệ Đơn Khiếu Nại / Tổng Đơn</div>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>項目 Chỉ số</th>
                            <th class="right">Shopee</th>
                            <th class="right">Tiktok</th>
                            <th class="right">Website</th>
                            <th class="right">Lazada</th>
                        </tr>
                    </thead>
                    <tbody id="wr-complaint-ratio-table"></tbody>
                </table>
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
const chartColors=["#8B6F47","#C9A961","#A89071","#6B8E5A","#B89970"];
const chartColorsAlt=["#3D2E1F","#8B6F47","#C9A961","#A6864B","#6B8E5A"];
const CSS_PRIMARY="#8B6F47";
const CSS_PRIMARY_RGBA="rgba(139,111,71,0.12)";
const CSS_TEXT_DARK="#3D2E1F";
const CSS_ACCENT="#C9A961";
const CSS_NEUTRAL="#A89071";
const platformNames={{shopee:"Shopee",tiktok:"TikTok",web:"Website",lazada:"Lazada"}};
const tabToKey={{shopee:"shopee",tiktok:"tiktok",lazada:"lazada",website:"web"}};
function resolveKey(tid){{return tabToKey[tid]||tid;}}
let currentMonth="{last_month}",charts={{}};
const DD={daily_str};
const PRODUCTS={products_str};
const WEEKLY={weekly_str};
const lookupFirstDate="{first_date}";
const lookupLastDate="{last_date}";
let lookupMode="single";
let lookupChannel="all";
let wrMode="single";
let wrChannel="all";

function fmt(n){{if(n>=1e9)return (n/1e9).toFixed(2)+"tỷ";if(n>=1e6)return (n/1e6).toFixed(1)+"tr";return n.toString().replace(/\\B(?=(\\d{{3}})+(?!\\d))/g,".");}}
function fmtFull(n){{return n.toString().replace(/\\B(?=(\\d{{3}})+(?!\\d))/g,".");}}
function chg(c,p){{if(!p)return{{pct:"N/A",arrow:"—",color:"#999"}};const pct=((c-p)/p*100).toFixed(1);return{{pct,arrow:pct>=0?"↑":"↓",color:pct>=0?"#6B8E5A":"#B5573D"}};}}
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
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"line",data:{{labels:ds,datasets:[{{label:"營收",data:dvs,borderColor:"#8B6F47",backgroundColor:"rgba(139,111,71,0.10)",pointBackgroundColor:"#8B6F47",pointBorderColor:"#FFFCF7",pointBorderWidth:2,pointRadius:4,fill:true,tension:.4}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
}}

/* ③ Trend + channel comparison */
function renderTrend(platforms,canvasId){{
    if(platforms.length===1){{
        const p=platforms[0];
        const data=allMonths.map(m=>getD(p,m).revenue);
        if(charts[canvasId])charts[canvasId].destroy();
        charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels:allMonths,datasets:[{{label:"營收",data:data,backgroundColor:"#8B6F47",borderRadius:6}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
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
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels,datasets:[{{label:"上月",data:prev,backgroundColor:"#A89071",borderRadius:6}},{{label:"本月",data:cur,backgroundColor:"#8B6F47",borderRadius:6}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
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
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels,datasets:[{{label:"上月",data:prev,backgroundColor:"#A89071",borderRadius:6}},{{label:"本月",data:cur,backgroundColor:"#8B6F47",borderRadius:6}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
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
    charts[canvasId]=new Chart(document.getElementById(canvasId),{{type:"bar",data:{{labels:allMonths,datasets:[{{label:"訂單",data:data,backgroundColor:"#C9A961",borderRadius:6}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
}}

/* Platform-specific fee for single platform */
function renderPlatformFees(platform,chartId,tableId){{
    const feesData=allMonths.map(m=>getD(platform,m).fees);
    if(charts[chartId])charts[chartId].destroy();
    charts[chartId]=new Chart(document.getElementById(chartId),{{type:"bar",data:{{labels:allMonths,datasets:[{{label:"平台費用",data:feesData,backgroundColor:"#A6864B",borderRadius:6}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
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
    charts["lookup-daily-chart"]=new Chart(document.getElementById("lookup-daily-chart"),{{type:"line",data:{{labels:dlabels,datasets:[{{label:"營收",data:dvs,borderColor:"#8B6F47",backgroundColor:"rgba(139,111,71,0.10)",pointBackgroundColor:"#8B6F47",pointBorderColor:"#FFFCF7",pointBorderWidth:2,pointRadius:4,fill:true,tension:.4}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});

    /* (3) Trend compare */
    if(charts["lookup-trend-chart"])charts["lookup-trend-chart"].destroy();
    charts["lookup-trend-chart"]=new Chart(document.getElementById("lookup-trend-chart"),{{type:"bar",data:{{labels:["週前","昨日","本期"],datasets:[{{label:"營收",data:[prevWeek.revenue,prev.revenue,cur.revenue],backgroundColor:["#B89970","#A89071","#8B6F47"],borderRadius:6}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});

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
    charts["lookup-mom-channel-chart"]=new Chart(document.getElementById("lookup-mom-channel-chart"),{{type:"bar",data:{{labels:chLabels,datasets:[{{label:"週前",data:chWeek,backgroundColor:"#B89970",borderRadius:6}},{{label:"昨日",data:chPrev,backgroundColor:"#A89071",borderRadius:6}},{{label:"本期",data:chCur,backgroundColor:"#8B6F47",borderRadius:6}}]}},options:{{responsive:true,maintainAspectRatio:false,scales:{{y:{{beginAtZero:true}}}}}}}});
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

/* ===== WEEKLY REPORT ===== */
function getWrRange(){{
    if(wrMode==="single"){{
        const v=document.getElementById("wr-date").value||lookupLastDate;
        return [v,v];
    }} else {{
        let s=document.getElementById("wr-start").value||lookupFirstDate;
        let e=document.getElementById("wr-end").value||lookupLastDate;
        if(s>e){{const t=s;s=e;e=t;}}
        return [s,e];
    }}
}}
function getWrPlatforms(){{
    return wrChannel==="all"?["shopee","tiktok","web","lazada"]:[wrChannel];
}}

function ymdToMonthKey(ymd){{ const p=ymd.split("-"); return "T"+parseInt(p[1],10); }}

function aggMonth(monthKey, platforms){{
    const result={{revenue:0,products:{{}}}};
    platforms.forEach(p=>{{
        if(!DD[p])return;
        Object.keys(DD[p]).forEach(d=>{{
            if(ymdToMonthKey(d)===monthKey){{
                const day=DD[p][d];
                result.revenue+=day.revenue;
                (day.products||[]).forEach(prod=>{{
                    if(!result.products[prod.name])result.products[prod.name]={{qty:0,revenue:0}};
                    result.products[prod.name].qty+=prod.qty;
                    result.products[prod.name].revenue+=prod.revenue;
                }});
            }}
        }});
    }});
    return result;
}}

function getCurrentWeekData(){{
    const cw=WEEKLY.current_week;
    if(cw&&WEEKLY.weeks&&WEEKLY.weeks[cw])return WEEKLY.weeks[cw];
    const keys=Object.keys(WEEKLY.weeks||{{}});
    return keys.length>0?WEEKLY.weeks[keys[0]]:null;
}}

function badgeDelta(cur,prev){{
    if(prev===null||prev===undefined||prev===0)return '<span class="badge neutral">—</span>';
    const delta=((cur-prev)/prev*100).toFixed(1);
    return delta>=0?`<span class="badge up">↑ ${{delta}}%</span>`:`<span class="badge down">↓ ${{Math.abs(delta).toFixed(1)}}%</span>`;
}}

function fmtCellNum(n){{ return (n===null||n===undefined)?"—":fmtFull(Math.round(n)); }}

function renderWeekly(){{
    const [s,e]=getWrRange();
    const platforms=getWrPlatforms();
    const cur=aggLookup(s,e,platforms);
    const [ws,we]=getPrevWeekRange(s,e);
    const prevWeek=aggLookup(ws,we,platforms);

    const same=(s===e);
    const periodLabel=same?fmtDate(s):`${{fmtDate(s)}} → ${{fmtDate(e)}}`;
    const prevPeriodLabel=(ws===we)?fmtDate(ws):`${{fmtDate(ws)}} → ${{fmtDate(we)}}`;
    const chLabel=wrChannel==="all"?"tất cả kênh":platformNames[wrChannel];
    document.getElementById("wr-compare-info").innerHTML=`Kỳ này: <b>${{periodLabel}}</b> | Tuần trước: ${{prevPeriodLabel}} | Kênh: <b>${{chLabel}}</b>`;
    document.getElementById("wr-revenue-meta").innerHTML=`Kỳ đang xem: <b>${{periodLabel}}</b> — Số liệu sàn lấy từ dashboard tự động; số liệu quyết toán nhập tay`;

    /* (1) Doanh thu sàn vs quyết toán */
    const wkData=getCurrentWeekData();
    const settlement=(wkData&&wkData.settlement)||{{}};
    const settMap={{shopee:settlement.shopee_settled||0,tiktok:settlement.tiktok_settled||0,lazada:settlement.lazada_settled||0,web:settlement.website_settled||0}};
    let revHtml="";
    let totRev=0,totSet=0;
    ["shopee","tiktok","lazada","web"].forEach(p=>{{
        const r=cur.byChannel[p]||0;
        const st=settMap[p]||0;
        const diff=r-st;
        const pct=r>0?((st/r)*100).toFixed(1)+"%":"—";
        revHtml+=`<tr><td><b>${{platformNames[p]}}</b></td><td class="right">${{fmtFull(r)}}</td><td class="right">${{fmtFull(st)}}</td><td class="right">${{fmtFull(diff)}}</td><td class="right">${{pct}}</td></tr>`;
        totRev+=r;totSet+=st;
    }});
    const totPct=totRev>0?((totSet/totRev)*100).toFixed(1)+"%":"—";
    revHtml+=`<tr style="font-weight:700;background:var(--bg-section)"><td>合計 Tổng Cộng</td><td class="right">${{fmtFull(totRev)}}</td><td class="right">${{fmtFull(totSet)}}</td><td class="right">${{fmtFull(totRev-totSet)}}</td><td class="right">${{totPct}}</td></tr>`;
    document.getElementById("wr-revenue-table").innerHTML=revHtml;

    /* (2) Top 10 SP Tuần + Tháng */
    const prodMap={{}};
    platforms.forEach(p=>{{
        if(!DD[p])return;
        Object.keys(DD[p]).forEach(d=>{{
            if(d>=s&&d<=e){{
                (DD[p][d].products||[]).forEach(prod=>{{
                    if(!prodMap[prod.name])prodMap[prod.name]={{qty:0,revenue:0}};
                    prodMap[prod.name].qty+=prod.qty;
                    prodMap[prod.name].revenue+=prod.revenue;
                }});
            }}
        }});
    }});
    const prevProdMap={{}};
    platforms.forEach(p=>{{
        if(!DD[p])return;
        Object.keys(DD[p]).forEach(d=>{{
            if(d>=ws&&d<=we){{
                (DD[p][d].products||[]).forEach(prod=>{{
                    if(!prevProdMap[prod.name])prevProdMap[prod.name]={{qty:0,revenue:0}};
                    prevProdMap[prod.name].qty+=prod.qty;
                    prevProdMap[prod.name].revenue+=prod.revenue;
                }});
            }}
        }});
    }});
    const topWeek=Object.entries(prodMap).sort((a,b)=>b[1].revenue-a[1].revenue).slice(0,10);
    let twH="";
    topWeek.forEach(([name,v],i)=>{{
        const prev=prevProdMap[name];
        const prevRev=prev?prev.revenue:0;
        twH+=`<tr><td><b>${{i+1}}</b></td><td>${{name}}</td><td class="right">${{fmtFull(v.qty)}}</td><td class="right">${{fmtFull(v.revenue)}}</td><td>${{badgeDelta(v.revenue,prevRev)}}</td></tr>`;
    }});
    if(!twH)twH='<tr><td colspan="5" class="wr-empty">Không có data SP trong khoảng đã chọn</td></tr>';
    document.getElementById("wr-top-week-table").innerHTML=twH;
    document.getElementById("wr-top-week-meta").textContent=periodLabel;

    const endMonth=ymdToMonthKey(e);
    const endMonthIdx=allMonths.indexOf(endMonth);
    const prevMonth=endMonthIdx>0?allMonths[endMonthIdx-1]:null;
    const monthData=aggMonth(endMonth,platforms);
    const prevMonthData=prevMonth?aggMonth(prevMonth,platforms):{{products:{{}}}};
    const topMonth=Object.entries(monthData.products).sort((a,b)=>b[1].revenue-a[1].revenue).slice(0,10);
    let tmH="";
    topMonth.forEach(([name,v],i)=>{{
        const prev=prevMonthData.products[name];
        const prevRev=prev?prev.revenue:0;
        tmH+=`<tr><td><b>${{i+1}}</b></td><td>${{name}}</td><td class="right">${{fmtFull(v.qty)}}</td><td class="right">${{fmtFull(v.revenue)}}</td><td>${{badgeDelta(v.revenue,prevRev)}}</td></tr>`;
    }});
    if(!tmH)tmH='<tr><td colspan="5" class="wr-empty">Không có data SP cho tháng</td></tr>';
    document.getElementById("wr-top-month-table").innerHTML=tmH;
    document.getElementById("wr-top-month-meta").textContent=endMonth+(prevMonth?` vs ${{prevMonth}}`:"");

    /* (3) Ads weekly */
    document.getElementById("wr-ad-week-label").textContent=wkData?wkData.label||"—":"Chưa có data";
    const adsW=(wkData&&wkData.ads_weekly)||{{shopee:{{}},tiktok:{{}}}};
    const sh=adsW.shopee||{{}},tk=adsW.tiktok||{{}};
    const rowsAdsW=[
        {{label:"Chi phí",sp:sh.prev_cost,sn:sh.cost,tp:tk.prev_cost,tn:tk.cost}},
        {{label:"Doanh số",sp:sh.prev_revenue,sn:sh.revenue,tp:tk.prev_revenue,tn:tk.revenue}},
        {{label:"Số đơn hàng",sp:sh.prev_orders,sn:sh.orders,tp:tk.prev_orders,tn:tk.orders}},
        {{label:"Sản phẩm bán",sp:sh.prev_products,sn:sh.products,tp:tk.prev_products,tn:tk.products}}
    ];
    let adwH="";
    rowsAdsW.forEach(r=>{{
        adwH+=`<tr><td><b>${{r.label}}</b></td><td class="right">${{fmtCellNum(r.sp)}}</td><td class="right">${{fmtCellNum(r.sn)}}</td><td>${{badgeDelta(r.sn||0,r.sp||0)}}</td><td class="right">${{fmtCellNum(r.tp)}}</td><td class="right">${{fmtCellNum(r.tn)}}</td><td>${{badgeDelta(r.tn||0,r.tp||0)}}</td></tr>`;
    }});
    document.getElementById("wr-ads-weekly-table").innerHTML=adwH;

    /* (4) Ads monthly */
    const adsM=(wkData&&wkData.ads_monthly)||{{shopee:{{}},tiktok:{{}}}};
    const shM=adsM.shopee||{{}},tkM=adsM.tiktok||{{}};
    const rowsAdsM=[
        {{label:"Chi phí",sc:shM.t_current_cost,spy:shM.t_prev_year_cost,sn:shM.t_next_cost,tc:tkM.t_current_cost,tpy:tkM.t_prev_year_cost,tn:tkM.t_next_cost}},
        {{label:"Doanh số",sc:shM.t_current_revenue,spy:shM.t_prev_year_revenue,sn:shM.t_next_revenue,tc:tkM.t_current_revenue,tpy:tkM.t_prev_year_revenue,tn:tkM.t_next_revenue}},
        {{label:"Số đơn hàng",sc:shM.t_current_orders,spy:shM.t_prev_year_orders,sn:shM.t_next_orders,tc:tkM.t_current_orders,tpy:tkM.t_prev_year_orders,tn:tkM.t_next_orders}},
        {{label:"Sản phẩm bán",sc:shM.t_current_products,spy:shM.t_prev_year_products,sn:shM.t_next_products,tc:tkM.t_current_products,tpy:tkM.t_prev_year_products,tn:tkM.t_next_products}}
    ];
    let admH="";
    rowsAdsM.forEach(r=>{{
        admH+=`<tr><td><b>${{r.label}}</b></td><td class="right">${{fmtCellNum(r.sc)}}</td><td class="right">${{fmtCellNum(r.spy)}}</td><td class="right">${{fmtCellNum(r.sn)}}</td><td class="right">${{fmtCellNum(r.tc)}}</td><td class="right">${{fmtCellNum(r.tpy)}}</td><td class="right">${{fmtCellNum(r.tn)}}</td></tr>`;
    }});
    document.getElementById("wr-ads-monthly-table").innerHTML=admH;

    /* (5) Ads ratio */
    const totalRev=(wkData&&wkData.ads_total_revenue)||{{}};
    function pct(a,b){{ return (b&&a)?((a/b)*100).toFixed(2)+"%":"—"; }}
    let ratH="";
    ratH+=`<tr><td>Chi phí</td><td class="right">${{fmtCellNum(shM.t_current_cost)}}</td><td class="right">${{fmtCellNum(shM.t_next_cost)}}</td><td class="right">${{fmtCellNum(tkM.t_current_cost)}}</td><td class="right">${{fmtCellNum(tkM.t_next_cost)}}</td></tr>`;
    ratH+=`<tr><td>Doanh số QC</td><td class="right">${{fmtCellNum(shM.t_current_revenue)}}</td><td class="right">${{fmtCellNum(shM.t_next_revenue)}}</td><td class="right">${{fmtCellNum(tkM.t_current_revenue)}}</td><td class="right">${{fmtCellNum(tkM.t_next_revenue)}}</td></tr>`;
    ratH+=`<tr><td>Tổng doanh thu bán hàng</td><td class="right">${{fmtCellNum(totalRev.shopee_t4)}}</td><td class="right">${{fmtCellNum(totalRev.shopee_t5)}}</td><td class="right">${{fmtCellNum(totalRev.tiktok_t4)}}</td><td class="right">${{fmtCellNum(totalRev.tiktok_t5)}}</td></tr>`;
    ratH+=`<tr><td><b>CP / Doanh số QC</b></td><td class="right"><b>${{pct(shM.t_current_cost,shM.t_current_revenue)}}</b></td><td class="right"><b>${{pct(shM.t_next_cost,shM.t_next_revenue)}}</b></td><td class="right"><b>${{pct(tkM.t_current_cost,tkM.t_current_revenue)}}</b></td><td class="right"><b>${{pct(tkM.t_next_cost,tkM.t_next_revenue)}}</b></td></tr>`;
    ratH+=`<tr style="background:var(--bg-section)"><td><b>CP / Tổng Doanh Thu</b></td><td class="right"><b>${{pct(shM.t_current_cost,totalRev.shopee_t4)}}</b></td><td class="right"><b>${{pct(shM.t_next_cost,totalRev.shopee_t5)}}</b></td><td class="right"><b>${{pct(tkM.t_current_cost,totalRev.tiktok_t4)}}</b></td><td class="right"><b>${{pct(tkM.t_next_cost,totalRev.tiktok_t5)}}</b></td></tr>`;
    document.getElementById("wr-ads-ratio-table").innerHTML=ratH;

    /* (6) Customer Care */
    const careRows=(wkData&&wkData.customer_care)||[];
    let carH="";
    careRows.forEach(r=>{{
        const conv=typeof r.conversion==="number"?r.conversion.toFixed(2)+"%":(r.conversion||"—");
        carH+=`<tr><td><b>${{r.channel}}</b></td><td class="right">${{fmtCellNum(r.chats)}}</td><td class="right">${{conv}}</td><td class="right">${{fmtCellNum(r.revenue)}}</td></tr>`;
    }});
    if(!carH)carH='<tr><td colspan="4" class="wr-empty">Chưa có data — cập nhật ở file weekly_report_data.json</td></tr>';
    document.getElementById("wr-care-table").innerHTML=carH;

    /* (7) Reviews */
    const reviewRows=(wkData&&wkData.reviews)||[];
    let revwH="";
    reviewRows.forEach(r=>{{
        revwH+=`<tr><td><b>${{r.channel}}</b></td><td class="right">${{r.rate5===null||r.rate5===undefined?"—":fmtFull(r.rate5)}}</td><td class="right">${{r.rate3===null||r.rate3===undefined?"—":fmtFull(r.rate3)}}</td><td class="right">${{r.rate1===null||r.rate1===undefined?"—":fmtFull(r.rate1)}}</td><td style="white-space:pre-line;font-size:0.88em">${{r.reason||"—"}}</td><td>${{r.status||"—"}}</td></tr>`;
    }});
    if(!revwH)revwH='<tr><td colspan="6" class="wr-empty">Chưa có data</td></tr>';
    document.getElementById("wr-reviews-table").innerHTML=revwH;

    /* (8) Complaints */
    const complaintRows=(wkData&&wkData.complaints)||[];
    let comH="";
    complaintRows.forEach(r=>{{
        comH+=`<tr><td><b>${{r.channel}}</b></td><td class="right">${{fmtCellNum(r.total_month)}}</td><td class="right">${{fmtCellNum(r.this_week)}}</td><td class="right">${{fmtCellNum(r.resolved)}}</td><td class="right">${{fmtCellNum(r.pending)}}</td><td style="white-space:pre-line;font-size:0.88em">${{r.reason||"—"}}</td></tr>`;
    }});
    if(!comH)comH='<tr><td colspan="6" class="wr-empty">Chưa có data</td></tr>';
    document.getElementById("wr-complaints-table").innerHTML=comH;

    /* (9) Complaint ratio */
    const cr=(wkData&&wkData.complaint_ratio)||{{}};
    function safePct(c,t){{ return t>0?((c/t)*100).toFixed(2)+"%":"—"; }}
    let crH="";
    crH+=`<tr><td>Số khiếu nại</td><td class="right">${{fmtCellNum((cr.shopee||{{}}).complaints)}}</td><td class="right">${{fmtCellNum((cr.tiktok||{{}}).complaints)}}</td><td class="right">${{fmtCellNum((cr.website||{{}}).complaints)}}</td><td class="right">${{fmtCellNum((cr.lazada||{{}}).complaints)}}</td></tr>`;
    crH+=`<tr><td>Tổng đơn hàng</td><td class="right">${{fmtCellNum((cr.shopee||{{}}).total_orders)}}</td><td class="right">${{fmtCellNum((cr.tiktok||{{}}).total_orders)}}</td><td class="right">${{fmtCellNum((cr.website||{{}}).total_orders)}}</td><td class="right">${{fmtCellNum((cr.lazada||{{}}).total_orders)}}</td></tr>`;
    crH+=`<tr style="background:var(--bg-section)"><td><b>Tỷ lệ KN / Tổng đơn</b></td><td class="right"><b>${{safePct((cr.shopee||{{}}).complaints,(cr.shopee||{{}}).total_orders)}}</b></td><td class="right"><b>${{safePct((cr.tiktok||{{}}).complaints,(cr.tiktok||{{}}).total_orders)}}</b></td><td class="right"><b>${{safePct((cr.website||{{}}).complaints,(cr.website||{{}}).total_orders)}}</b></td><td class="right"><b>${{safePct((cr.lazada||{{}}).complaints,(cr.lazada||{{}}).total_orders)}}</b></td></tr>`;
    document.getElementById("wr-complaint-ratio-table").innerHTML=crH;
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
    else if(tid==="weekly")renderWeekly();
    else platform(resolveKey(tid));
}});}});

/* Weekly report control handlers */
document.querySelectorAll('#weekly .mode-btn[data-target="wr"]').forEach(b=>{{b.addEventListener("click",e=>{{
    wrMode=e.target.dataset.mode;
    document.querySelectorAll('#weekly .mode-btn[data-target="wr"]').forEach(t=>t.classList.remove("active"));
    e.target.classList.add("active");
    document.getElementById("wr-single-controls").style.display = wrMode==="single"?"":"none";
    document.getElementById("wr-range-controls").style.display = wrMode==="range"?"":"none";
    renderWeekly();
}});}});
document.querySelectorAll('#weekly .ch-pill[data-target="wr"]').forEach(b=>{{b.addEventListener("click",e=>{{
    wrChannel=e.target.dataset.channel;
    document.querySelectorAll('#weekly .ch-pill[data-target="wr"]').forEach(t=>t.classList.remove("active"));
    e.target.classList.add("active");
    renderWeekly();
}});}});
["wr-date","wr-start","wr-end"].forEach(id=>{{
    const el=document.getElementById(id);
    if(el)el.addEventListener("change",renderWeekly);
}});

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
    output_path = sys.argv[1] if len(sys.argv) > 1 else "/sessions/relaxed-laughing-hamilton/mnt/Loho_Dashboard/Dashboard_Loho_House_2026.html"

    print("=== Loho House Dashboard Updater v2 ===")
    print(f"Output: {output_path}")
    print()

    print("Downloading 8 sheets...")
    csv_paths = {}
    for name, info in SHEETS.items():
        print(f"  -> {name} (gid={info['gid']})...", end=" ")
        path = download_csv(name, info["gid"])
        if path:
            csv_paths[name] = path
            lines = sum(1 for _ in open(path, "r", encoding="utf-8", errors="replace"))
            print(f"OK ({lines} lines)")
        else:
            print("FAILED")

    print("\nProcessing revenue data...")
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

    print("\nProcessing categories...")
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

    print("\nEnriching products from raw sheets...")
    for platform in ["web", "lazada"]:
        raw_name = platform + "_raw"
        if raw_name in csv_paths:
            enrich_products_from_raw(platform, csv_paths[raw_name], platforms_data, daily_data)

    print("\nBuilding data JSON...")
    data_json = build_data_json(platforms_data, categories_data)
    daily_json = build_daily_json(daily_data, daily_categories_data)
    products_json = build_products_index(daily_data)
    print(f"  Built products index: {len(products_json)} unique products")

    print("\nLoading weekly report manual data...")
    weekly_data = load_weekly_report_data()
    n_weeks = len(weekly_data.get("weeks", {}))
    print(f"  Loaded {n_weeks} week(s) of manual data")

    print(f"\nGenerating HTML -> {output_path}")
    generate_html(data_json, daily_json, products_json, output_path, weekly_data=weekly_data)

    file_size = os.path.getsize(output_path)
    print(f"\nDone! File size: {file_size:,} bytes")
    print(f"Updated: {(datetime.now(VN_TZ) if VN_TZ else datetime.now()).strftime('%d/%m/%Y %H:%M')}")


if __name__ == "__main__":
    main()
