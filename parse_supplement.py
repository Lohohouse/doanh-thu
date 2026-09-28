# -*- coding: utf-8 -*-
"""
parse_supplement.py — Đọc file gốc (Shopee/TikTok) -> monthly_supplement.json
Nguồn cho dashboard: đơn hoàn, phí sàn thực + chi tiết (D3), DT thực, DT theo nguồn (B4), đơn hủy.
Chạy: python parse_supplement.py <thư_mục_chứa_file> [output.json]
File cần (mỗi tháng Tx):
  - SHOPEE Doanh thu - CP Tx.pdf        (phí sàn, đơn hoàn, DT thực)
  - TIKTOK Doanh thu - CP Tx.xlsx       (sheet 'Báo cáo', giá trị cột F)
  - Shopee doanh số Tx.xlsx             (nguồn B4 + đơn hủy)
  - Tiktok doanh số Tx.xlsx             (GMV, hoàn tiền, đơn)
"""
import sys, os, re, glob, json

try:
    import openpyxl
except ImportError:
    os.system(sys.executable + " -m pip install openpyxl --break-system-packages -q")
    import openpyxl
try:
    import pdfplumber
except ImportError:
    os.system(sys.executable + " -m pip install pdfplumber --break-system-packages -q")
    import pdfplumber


def _num(s):
    """'-25,328,547' / '₫876,710,277' / '1.234' -> float. Trả None nếu không parse được."""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = str(s).replace("₫", "").replace("VND", "").strip()
    neg = t.startswith("-") or t.startswith("−")
    t = re.sub(r"[^\d]", "", t)          # bỏ mọi dấu phân tách (, . khoảng trắng)
    if t == "":
        return None
    v = float(t)
    return -v if neg else v


def find_file(folder, *patterns):
    for p in patterns:
        g = glob.glob(os.path.join(folder, p))
        if g:
            return g[0]
    return None


# ---------------- SHOPEE CP PDF ----------------
def parse_shopee_cp(path):
    """Đọc block 'Tổng kết thanh toán đã chuyển' trong PDF Shopee."""
    with pdfplumber.open(path) as pdf:
        txt = pdf.pages[0].extract_text() or ""
    cut = txt.find("Thông tin thanh toán")
    head = txt[:cut] if cut > 0 else txt
    def grab(label):
        # số cuối cùng trên dòng chứa label
        for line in head.split("\n"):
            if label.lower() in line.lower():
                m = re.findall(r"[₫\-−]?\s?[\d.,]{2,}", line)
                if m:
                    return _num(m[-1])
        return None
    gia_sp   = grab("Giá sản phẩm")
    hoan     = grab("Số tiền hoàn lại")
    giamgia  = grab("Mã ưu đãi do Người Bán chịu")
    vc_net   = grab("Phí vận chuyển (không tính trợ giá)")
    phi_gd   = grab("Phí giao dịch")
    f_codinh = grab("Phí cố định")
    f_dv     = grab("Phí Dịch Vụ")
    f_xlgd   = grab("Phí xử lý giao dịch")
    f_hhtt   = grab("Phí hoa hồng Tiếp thị liên kết")
    f_piship = grab("Phí dịch vụ PiShip")
    da_chuyen= grab("Tổng thanh toán đã chuyển")
    # phí sàn = phí giao dịch (gồm 5 khoản con) + phí vận chuyển ròng không trợ giá
    phi_san = abs(phi_gd or 0) + abs(vc_net or 0)
    chi_tiet = []
    for ten, gt in [("Phí cố định", f_codinh), ("Phí Dịch Vụ", f_dv),
                    ("Phí xử lý giao dịch", f_xlgd), ("Phí hoa hồng Tiếp thị liên kết", f_hhtt),
                    ("Phí dịch vụ PiShip", f_piship), ("Phí vận chuyển ròng", vc_net)]:
        if gt:
            chi_tiet.append({"ten": ten, "gia_tri": abs(gt)})
    return {
        "gia_sp": gia_sp,
        "don_hoan": abs(hoan) if hoan else 0,     # số tiền hoàn
        "giam_gia": abs(giamgia) if giamgia else 0,
        "phi_san": phi_san,
        "dt_thuc": da_chuyen,                      # tổng thanh toán đã chuyển
        "chi_tiet_phi": chi_tiet,
    }


# ---------------- TIKTOK CP xlsx (sheet 'Báo cáo', giá trị cột F=idx5) ----------------
def parse_tiktok_cp(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Báo cáo" not in wb.sheetnames:
        return {}
    ws = wb["Báo cáo"]
    rows = list(ws.iter_rows(values_only=True))
    def val_for(label, exact=False):
        for r in rows:
            lab = " ".join(str(x) for x in r[:5] if isinstance(x, str))
            if (label == lab.strip()) if exact else (label.lower() in lab.lower()):
                # giá trị = số đầu tiên từ cột E trở đi
                for c in r[4:]:
                    if isinstance(c, (int, float)):
                        return float(c)
                    n = _num(c)
                    if n is not None and str(c).strip() not in ("VND", "UTC+7"):
                        return n
        return None
    tong_qt   = val_for("Tổng số tiền quyết toán")
    tong_dt   = val_for("Tổng doanh thu")
    hoan      = val_for("khoản hoàn tiền")            # Tổng phụ của khoản hoàn tiền
    tong_phi  = val_for("Tổng phí")
    dieu_chinh= val_for("Điều chỉnh")
    # chi tiết phí chính
    chi_tiet = []
    for lbl in ["Phí giao dịch", "Phí hoa hồng của TikTok Shop", "Phí vận chuyển của người bán",
                "Hoa hồng liên kết", "Phí quảng cáo GMV Max"]:
        v = val_for(lbl)
        if v:
            chi_tiet.append({"ten": lbl, "gia_tri": abs(v)})
    return {
        "dt_qt": tong_qt,
        "tong_dt": tong_dt,
        "don_hoan": abs(hoan) if hoan else 0,
        "phi_san": abs(tong_phi) if tong_phi else 0,
        "dieu_chinh": dieu_chinh,
        "dt_thuc": tong_qt,               # tổng số tiền quyết toán = thực nhận
        "chi_tiet_phi": chi_tiet,
    }


# ---------------- SHOPEE doanh số (nguồn B4 + đơn hủy) ----------------
def parse_shopee_ds(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    # đơn hủy + tổng DS từ sheet 'Đơn hàng đã đặt'
    for sn in wb.sheetnames:
        if sn.strip().startswith("Đơn hàng đã đặt"):
            r = list(wb[sn].iter_rows(values_only=True))
            if len(r) > 1:
                out["tong_ds"] = _num(r[1][1])
                out["so_don"]  = _num(r[1][3])
                out["don_huy"] = _num(r[1][8])
                out["ds_huy"]  = _num(r[1][9])
            break
    # nguồn: sheet 'Nguồn truy cập cho Đơn hàng...' hàng 'Đơn hàng đã đặt'
    for sn in wb.sheetnames:
        if sn.startswith("Nguồn truy cập cho Đơn hàng"):
            for r in wb[sn].iter_rows(values_only=True):
                if r and str(r[1] or "").strip().startswith("Đơn hàng đã đặt"):
                    out["nguon"] = {
                        "the_sp":  _num(r[3]),
                        "live":    _num(r[4]),
                        "video":   _num(r[5]),
                        "ttlk":    _num(r[6]),   # đối tác/tiếp thị liên kết
                        "quang_cao": _num(r[7]),
                    }
                    break
            break
    return out


# ---------------- TIKTOK doanh số (GMV, hoàn, đơn) ----------------
def parse_tiktok_ds(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    hdr = None; val = None
    for i, r in enumerate(rows):
        cells = [str(x) if x is not None else "" for x in r]
        if "GMV" in cells and "Đơn hàng" in cells:
            hdr = cells
            if i + 1 < len(rows):
                val = rows[i + 1]
            break
    out = {}
    if hdr and val:
        idx = {h: k for k, h in enumerate(hdr)}
        def g(name):
            return _num(val[idx[name]]) if name in idx else None
        out = {"gmv": g("GMV"), "so_don": g("Đơn hàng"),
               "hoan_tien": g("Hoàn tiền"), "tong_dt": g("Tổng doanh thu"),
               "so_mon": g("Số món bán ra")}
    return out


def parse_month(folder, tag):
    rec = {}
    f = find_file(folder, f"SHOPEE Doanh thu - CP {tag}.pdf", f"*SHOPEE*CP*{tag}.pdf")
    sc = parse_shopee_cp(f) if f else {}
    f = find_file(folder, f"Shopee doanh số {tag}.xlsx", f"Shopee doanh s*{tag}.xlsx")
    sd = parse_shopee_ds(f) if f else {}
    f = find_file(folder, f"TIKTOK Doanh thu - CP {tag}.xlsx", f"*TIKTOK*CP*{tag}.xlsx")
    tc = parse_tiktok_cp(f) if f else {}
    f = find_file(folder, f"Tiktok doanh số {tag}.xlsx", f"Tiktok doanh s*{tag}.xlsx")
    td = parse_tiktok_ds(f) if f else {}
    if sc or sd:
        rec["shopee"] = {**sc, **sd}
    if tc or td:
        rec["tiktok"] = {**tc, **td}
    return rec


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    out_path = sys.argv[2] if len(sys.argv) > 2 else "monthly_supplement.json"
    data = {"nam": 2026, "cap_nhat": None, "thang": {}}
    for m in range(1, 13):
        tag = f"T{m}"
        if not (find_file(folder, f"*CP {tag}.pdf", f"*CP {tag}.xlsx") or
                find_file(folder, f"*doanh s* {tag}.xlsx")):
            continue
        rec = parse_month(folder, tag)
        if rec:
            data["thang"][tag] = rec
            print(f"[{tag}] OK  kênh: {list(rec.keys())}")
    import datetime
    data["cap_nhat"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    with open(out_path, "w", encoding="utf-8") as fo:
        json.dump(data, fo, ensure_ascii=False, indent=1)
    print("WROTE", out_path, "| thang:", list(data["thang"].keys()))


if __name__ == "__main__":
    main()
