"""
ỨNG DỤNG QUẢN LÝ KHO HÀNG
--------------------------
Chạy bằng lệnh:  streamlit run quan_ly_kho.py

Cấu trúc kho:
- KỆ HÀNG DO NGƯỜI DÙNG TỰ TẠO (chức năng "Tạo kệ hàng"): nhập SỐ HÀNG (tầng) và SỐ CỘT (ô mỗi tầng).
  Có thể tạo nhiều kệ (Kệ 1, Kệ 2, ...), mỗi kệ có kích thước riêng.
  CHƯA CÓ KỆ NÀO thì chưa dùng được Nhập hàng / Xuất hàng.
- Có thể XÓA kệ, nhưng chỉ khi kệ đang trống (không còn sản phẩm nào).
- Số ô của 1 kệ = số hàng x số cột
- Số VỊ TRÍ MỖI Ô do người dùng nhập khi tạo kệ (mặc định 5); mọi ô trong cùng 1 kệ có cùng số vị trí.
- MỖI VỊ TRÍ CHỈ CHỨA ĐÚNG 1 SẢN PHẨM (1 mã sản phẩm duy nhất) — không dán nhiều mã
  vào cùng 1 vị trí.
- Do đó 1 ô chứa tối đa bằng số vị trí của kệ đó (có thể nhiều tên khác nhau, hoặc cùng tên nhưng
  khác loại/nhiều sản phẩm cùng loại, miễn tổng số sản phẩm không vượt quá số vị trí).

Cú pháp mã sản phẩm:
    {Mã ô}-{TÊN SP}-{LOẠI}-{STT}
    STT được đếm riêng theo từng cặp (Tên sản phẩm, Loại) trong CÙNG 1 Ô,
    và được đánh lại từ 01 khi đổi sang loại khác.

    Ví dụ: trong ô K1T101, "Áo thun" có 2 loại, Cotton chiếm 3 vị trí, Da Bò chiếm 2 vị trí:
        Vị trí 1: K1T101-AOTHUN-COTTON-01
        Vị trí 2: K1T101-AOTHUN-COTTON-02
        Vị trí 3: K1T101-AOTHUN-COTTON-03
        Vị trí 4: K1T101-AOTHUN-DABO-01
        Vị trí 5: K1T101-AOTHUN-DABO-02

- Mã ô: K{kệ}T{tầng}{ô}   (VD: K1T101)
- Giao diện đồng bộ: nền xám nhạt (tự đổi tông tối), tiêu đề gradient, thẻ số liệu, nút bo góc, tab/thông báo/form cùng phong cách.
- 4 chức năng chính: Kiểm tra - Nhập - Xuất - Lịch sử (xem/lọc/tải CSV lịch sử nhập, xuất)
- TẠO MÃ QR hàng loạt (Kệ-Tên-Loại-STT) và xuất ra file Word để in tem. Cần: pip install python-docx qrcode
- KIỂM TRA: quét QR vị trí ô để xem trong ô đó có những sản phẩm nào, hoặc xem sơ đồ kho.
- NHẬP / XUẤT có 2 cách (2 tab): quét QR bằng camera, hoặc nhập / xuất TAY chọn ô trên sơ đồ kệ.
- NHẬP HÀNG bằng quét QR: (1) quét QR ô (VD: K1T101), (2) quét QR sản phẩm có nội dung
  dạng  Kệ-Tên-Loại-STT  (VD: K1-VFD-3P-01). Sản phẩm tự vào vị trí trống đầu tiên.
- XUẤT HÀNG bằng quét QR: chỉ cần quét QR sản phẩm (Kệ-Tên-Loại-STT), app tự tìm vị trí và xuất.
  Camera quét TRỰC TIẾP (khung vuông), tự nhận diện khi đưa QR vào khung, không cần bấm chụp.
  Chỉ cần: pip install streamlit  (không cần thư viện quét QR nào khác)
- Dữ liệu được lưu vào CƠ SỞ DỮ LIỆU SQL (SQLite, file kho.db) để không bị mất khi tắt/mở lại app
- Phải ĐĂNG NHẬP mới dùng được app; chưa có tài khoản thì tạo ngay trên màn hình đăng nhập.
  Tài khoản lưu trong bảng tai_khoan của kho.db (mật khẩu được băm PBKDF2 + muối, không lưu dạng chữ thường).
  Lần chạy đầu tiên, nếu còn các file JSON cũ (tai_khoan.json, du_lieu_tai_khoan/, phien_dang_nhap.json)
  thì dữ liệu được tự động chuyển sang SQL (file JSON cũ giữ nguyên làm bản sao lưu).
  Tên tài khoản PHÂN BIỆT chữ hoa / chữ thường (An_Kho và an_kho là 2 tài khoản khác nhau, mỗi tài khoản 1 kho riêng).
"""

import io
import math
import pandas as pd
import streamlit as st
import json
import sqlite3
from contextlib import contextmanager
import hashlib
import hmac
import secrets
import time
import os
import re
import unicodedata
from html import escape as _esc
from datetime import datetime

import tempfile
from pathlib import Path
import streamlit.components.v1 as components

# ============ CẤU HÌNH CHUNG ============
MAX_HANG = 10        # số hàng (tầng) tối đa khi tạo 1 kệ (có thể chỉnh)
MAX_COT = 10         # số cột (ô mỗi tầng) tối đa khi tạo 1 kệ (tối đa 99 vì mã ô dùng 2 chữ số)
SO_VI_TRI = 5        # số vị trí MẶC ĐỊNH trong 1 ô khi tạo kệ (mỗi vị trí = 1 sản phẩm)
MAX_VI_TRI = 20      # số vị trí tối đa trong 1 ô khi tạo kệ (có thể chỉnh)
FILE_DB = "kho.db"                      # cơ sở dữ liệu SQLite chứa tài khoản, kho, lịch sử, phiên đăng nhập
# --- Các file JSON CŨ: chỉ còn dùng để chuyển dữ liệu sang SQL ở lần chạy đầu ---
FILE_DATA = "kho_data.json"
FILE_LICH_SU = "lich_su_kho.json"
THU_MUC_DU_LIEU = "du_lieu_tai_khoan"
FILE_TAI_KHOAN = "tai_khoan.json"
MK_TOI_THIEU = 6       # độ dài mật khẩu tối thiểu
SAI_TOI_DA = 5         # sai mật khẩu quá số lần này thì tạm khóa đăng nhập
KHOA_GIAY = 30         # số giây tạm khóa
FILE_PHIEN = "phien_dang_nhap.json"   # (cũ) phiên đăng nhập; nay lưu trong bảng phien (chỉ lưu bản băm của mã phiên)
COOKIE_PHIEN = "kho_phien"            # tên cookie trên trình duyệt giữ mã phiên
PHIEN_NGAY = 30                       # nhớ đăng nhập bao nhiêu ngày

st.set_page_config(page_title="Quản lý kho hàng", page_icon="", layout="wide")


# ============ CHUẨN HÓA CHUỖI CHO MÃ SẢN PHẨM ============
def chuan_hoa(chuoi):
    """Bỏ dấu tiếng Việt, bỏ khoảng trắng/ký tự đặc biệt, viết HOA -> dùng để ghép mã."""
    chuoi = chuoi.strip()
    chuoi = unicodedata.normalize("NFD", chuoi)
    chuoi = "".join(ch for ch in chuoi if unicodedata.category(ch) != "Mn")
    chuoi = chuoi.replace("đ", "d").replace("Đ", "D")
    chuoi = re.sub(r"[^A-Za-z0-9]", "", chuoi)
    return chuoi.upper()


# ============ HÀM XỬ LÝ DỮ LIỆU ============
def tao_kho_rong():
    """
    Mỗi ô là 1 danh sách có ĐÚNG số-vị-trí-của-kệ phần tử (len(ô)).
    Mỗi phần tử (1 vị trí): None (trống) hoặc:
    {
        "ten_hang": str,
        "loai": str,
        "stt": int
    }
    -> mỗi vị trí chỉ chứa DUY NHẤT 1 sản phẩm / 1 mã.
    """
    return {}      # kho mới chưa có kệ nào — kệ được thêm bằng chức năng "Tạo kệ hàng"


def tao_ke_rong(so_hang, so_cot, so_vi_tri=SO_VI_TRI):
    """Một kệ trống: `so_hang` tầng, mỗi tầng `so_cot` ô, mỗi ô `so_vi_tri` vị trí trống (None = trống)."""
    return [[[None for _ in range(so_vi_tri)] for _ in range(so_cot)] for _ in range(so_hang)]


# ============ CƠ SỞ DỮ LIỆU SQL (SQLite) ============
SCHEMA = """
CREATE TABLE IF NOT EXISTS tai_khoan (
    ten       TEXT PRIMARY KEY,            -- phân biệt chữ hoa / chữ thường (so sánh nhị phân)
    muoi      TEXT NOT NULL,
    mat_khau  TEXT NOT NULL,               -- PBKDF2-SHA256 + muối
    ngay_tao  TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS ke (
    tai_khoan TEXT NOT NULL REFERENCES tai_khoan(ten) ON DELETE CASCADE,
    ten_ke    TEXT NOT NULL,
    so_hang   INTEGER NOT NULL,
    so_cot    INTEGER NOT NULL,
    so_vi_tri INTEGER NOT NULL,
    PRIMARY KEY (tai_khoan, ten_ke)
);
CREATE TABLE IF NOT EXISTS vi_tri (        -- chỉ lưu các vị trí ĐANG CÓ HÀNG; vị trí trống = không có dòng
    tai_khoan TEXT NOT NULL,
    ten_ke    TEXT NOT NULL,
    hang      INTEGER NOT NULL,
    cot       INTEGER NOT NULL,
    vt        INTEGER NOT NULL,
    ten_hang  TEXT NOT NULL,
    loai      TEXT NOT NULL,
    stt       INTEGER NOT NULL,
    PRIMARY KEY (tai_khoan, ten_ke, hang, cot, vt),
    FOREIGN KEY (tai_khoan, ten_ke) REFERENCES ke(tai_khoan, ten_ke) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS lich_su (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    tai_khoan   TEXT NOT NULL REFERENCES tai_khoan(ten) ON DELETE CASCADE,
    thoi_gian   TEXT NOT NULL,
    hanh_dong   TEXT NOT NULL,
    ma_san_pham TEXT,
    ten_hang    TEXT,
    loai        TEXT
);
CREATE INDEX IF NOT EXISTS idx_lich_su_tk ON lich_su(tai_khoan, id);
CREATE TABLE IF NOT EXISTS phien (
    ma_bam  TEXT PRIMARY KEY,              -- SHA-256 của mã phiên (mã gốc chỉ nằm trong cookie)
    ten     TEXT NOT NULL REFERENCES tai_khoan(ten) ON DELETE CASCADE,
    het_han REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS meta (khoa TEXT PRIMARY KEY, gia_tri TEXT);
"""

# Bảng thêm sau này: tách riêng để luôn tự tạo được cả khi app đang chạy với CSDL / bộ nhớ đệm của bản cũ
SCHEMA_TEN_SAN_PHAM = """
CREATE TABLE IF NOT EXISTS ten_san_pham (  -- mã sản phẩm (dùng trong mã QR) -> tên sản phẩm (hiển thị)
    tai_khoan TEXT NOT NULL REFERENCES tai_khoan(ten) ON DELETE CASCADE,
    viet_tat  TEXT NOT NULL,                -- đã chuẩn hóa (viết hoa, bỏ dấu, bỏ khoảng trắng) như trong mã
    day_du    TEXT NOT NULL,
    PRIMARY KEY (tai_khoan, viet_tat)
);
CREATE TABLE IF NOT EXISTS loai_san_pham (  -- (mã sản phẩm, mã loại) dùng trong mã QR -> tên loại (hiển thị)
    tai_khoan TEXT NOT NULL REFERENCES tai_khoan(ten) ON DELETE CASCADE,
    ma_sp     TEXT NOT NULL,                -- mã sản phẩm đã chuẩn hóa
    ma_loai   TEXT NOT NULL,                -- mã loại đã chuẩn hóa
    ten_loai  TEXT NOT NULL,
    PRIMARY KEY (tai_khoan, ma_sp, ma_loai)
);
"""
SCHEMA = SCHEMA + SCHEMA_TEN_SAN_PHAM
PHIEN_BAN_SCHEMA = 3     # tăng số này khi thêm bảng mới để app chạy lại bước tạo bảng



def _doc_json(duong_dan, mac_dinh):
    try:
        with open(duong_dan, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return mac_dinh


def _chuyen_json_sang_sql(con):
    """Chạy 1 lần: nạp dữ liệu từ các file JSON cũ (nếu có) vào SQL. File JSON cũ được giữ nguyên làm bản sao lưu."""
    ds = _doc_json(FILE_TAI_KHOAN, {})
    if not isinstance(ds, dict):
        ds = {}
    ten_dau = None
    for khoa, tk in ds.items():
        if not isinstance(tk, dict):
            continue
        thu_muc = tk.get("thu_muc") or ("tk_" + khoa)         # tài khoản đời cũ chưa có "thu_muc"
        ten = tk.get("hien_thi", khoa)
        if not tk.get("muoi") or not tk.get("mat_khau"):
            continue
        ten_dau = ten_dau or ten
        con.execute("INSERT OR IGNORE INTO tai_khoan(ten, muoi, mat_khau, ngay_tao) VALUES (?,?,?,?)",
                    (ten, tk["muoi"], tk["mat_khau"], tk.get("ngay_tao") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        p_kho = Path(THU_MUC_DU_LIEU) / thu_muc / FILE_DATA
        p_ls = Path(THU_MUC_DU_LIEU) / thu_muc / FILE_LICH_SU
        if ten == ten_dau:                                      # file dùng chung đời cũ -> tài khoản tạo sớm nhất
            if not p_kho.exists() and Path(FILE_DATA).exists():
                p_kho = Path(FILE_DATA)
            if not p_ls.exists() and Path(FILE_LICH_SU).exists():
                p_ls = Path(FILE_LICH_SU)
        kho = _doc_json(p_kho, {})
        if isinstance(kho, dict):
            _ghi_kho(con, ten, kho)
        ls = _doc_json(p_ls, [])
        if isinstance(ls, list):
            con.executemany(
                "INSERT INTO lich_su(tai_khoan, thoi_gian, hanh_dong, ma_san_pham, ten_hang, loai) VALUES (?,?,?,?,?,?)",
                [(ten, m.get("thoi_gian", ""), m.get("hanh_dong", ""), m.get("ma_san_pham"), m.get("ten_hang"), m.get("loai"))
                 for m in ls if isinstance(m, dict)])
    phien = _doc_json(FILE_PHIEN, {})
    if isinstance(phien, dict):
        for ma_bam, m in phien.items():
            if isinstance(m, dict) and m.get("ten") and m.get("het_han", 0) > time.time():
                con.execute("INSERT OR IGNORE INTO phien(ma_bam, ten, het_han) "
                            "SELECT ?, ?, ? WHERE EXISTS (SELECT 1 FROM tai_khoan WHERE ten = ?)",
                            (ma_bam, m["ten"], m["het_han"], m["ten"]))


@st.cache_resource
def _khoi_tao_db(phien_ban=PHIEN_BAN_SCHEMA):
    """Tạo bảng (nếu chưa có) và chuyển dữ liệu JSON cũ sang SQL đúng 1 lần."""
    con = sqlite3.connect(FILE_DB, timeout=30)
    try:
        con.execute("PRAGMA journal_mode = WAL")          # đọc / ghi đồng thời tốt hơn khi nhiều người dùng
        con.execute("PRAGMA foreign_keys = ON")
        with con:
            con.executescript(SCHEMA)
            da_chuyen = con.execute("SELECT 1 FROM meta WHERE khoa = 'da_chuyen_json'").fetchone()
            if not da_chuyen:
                _chuyen_json_sang_sql(con)
                con.execute("INSERT INTO meta(khoa, gia_tri) VALUES ('da_chuyen_json', '1')")
    finally:
        con.close()
    return True


@contextmanager
def _db():
    """Mở kết nối SQLite; thoát khối `with` thì COMMIT (hoặc ROLLBACK nếu có lỗi) rồi đóng kết nối."""
    _khoi_tao_db()
    con = sqlite3.connect(FILE_DB, timeout=30)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    try:
        with con:
            yield con
    finally:
        con.close()


# ---------- Tài khoản ----------
def lay_tai_khoan(ten):
    """Trả về {"ten", "muoi", "mat_khau", "ngay_tao"} hoặc None. Tên phải ĐÚNG chữ hoa / chữ thường."""
    with _db() as con:
        r = con.execute("SELECT ten, muoi, mat_khau, ngay_tao FROM tai_khoan WHERE ten = ?", (ten,)).fetchone()
    return dict(r) if r else None


def them_tai_khoan(ten, muoi, mat_khau_bam):
    """Thêm tài khoản mới. Trả về False nếu tên đã tồn tại."""
    try:
        with _db() as con:
            con.execute("INSERT INTO tai_khoan(ten, muoi, mat_khau, ngay_tao) VALUES (?,?,?,?)",
                        (ten, muoi, mat_khau_bam, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        return True
    except sqlite3.IntegrityError:
        return False


def _ten_khoa(ten=None):
    """Tên tài khoản ĐÚNG chữ hoa/thường (lấy từ tham số hoặc từ phiên đăng nhập) và phải đang tồn tại trong SQL."""
    ten = (ten or st.session_state.get("nguoi_dung") or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9_.\-]{3,30}", ten):
        raise PermissionError("Chưa đăng nhập: không có kho dữ liệu để đọc / ghi.")
    if lay_tai_khoan(ten) is None:
        raise PermissionError("Tài khoản không tồn tại.")
    return ten


# ---------- Kho hàng ----------
def _ghi_kho(con, ten, kho):
    """Ghi toàn bộ kho của tài khoản `ten` (xóa cũ, chèn mới) trong 1 giao dịch của kết nối `con`."""
    con.execute("DELETE FROM ke WHERE tai_khoan = ?", (ten,))          # ON DELETE CASCADE xóa luôn vi_tri
    for ten_ke, ke in kho.items():
        so_hang = len(ke)
        so_cot = len(ke[0]) if so_hang else 0
        so_vt = len(ke[0][0]) if so_cot else 0
        con.execute("INSERT INTO ke(tai_khoan, ten_ke, so_hang, so_cot, so_vi_tri) VALUES (?,?,?,?,?)",
                    (ten, ten_ke, so_hang, so_cot, so_vt))
        con.executemany(
            "INSERT INTO vi_tri(tai_khoan, ten_ke, hang, cot, vt, ten_hang, loai, stt) VALUES (?,?,?,?,?,?,?,?)",
            [(ten, ten_ke, h, c, i, v["ten_hang"], v["loai"], int(v["stt"]))
             for h, hang in enumerate(ke) for c, o in enumerate(hang) for i, v in enumerate(o) if v is not None])


def doc_du_lieu(ten=None):
    """Đọc kho của tài khoản từ SQL, trả về đúng cấu trúc cũ: {tên_kệ: [tầng][ô][vị trí] = None | {"ten_hang","loai","stt"}}."""
    ten = _ten_khoa(ten)
    kho = tao_kho_rong()
    with _db() as con:
        for r in con.execute("SELECT ten_ke, so_hang, so_cot, so_vi_tri FROM ke WHERE tai_khoan = ? ORDER BY rowid", (ten,)):
            kho[r["ten_ke"]] = tao_ke_rong(r["so_hang"], r["so_cot"], r["so_vi_tri"])
        for r in con.execute("SELECT ten_ke, hang, cot, vt, ten_hang, loai, stt FROM vi_tri WHERE tai_khoan = ?", (ten,)):
            kho[r["ten_ke"]][r["hang"]][r["cot"]][r["vt"]] = {"ten_hang": r["ten_hang"], "loai": r["loai"], "stt": r["stt"]}
    return kho


def luu_du_lieu(kho):
    ten = _ten_khoa()
    with _db() as con:                       # 1 giao dịch: hoặc ghi đủ, hoặc không ghi gì (không bao giờ hỏng dở)
        _ghi_kho(con, ten, kho)


def _dinh_dang_stt_trong_ma(ma):
    """Chuẩn hóa mã sản phẩm trong lịch sử: STT cuối về 2 chữ số (...-001 -> ...-01, từ 100 trở lên giữ nguyên)
    và mã cũ có mã ô ở đầu (K1T101-VFD-3P-01) đổi sang dạng mới theo kệ (K1-VFD-3P-01)."""
    ma = re.sub(r"^(K\d+)T\d+-", r"\1-", str(ma))
    return re.sub(r"-(\d+)$", lambda m: f"-{int(m.group(1)):02d}", ma)


# ---------- Lịch sử nhập / xuất ----------
def doc_lich_su():
    """Đọc toàn bộ lịch sử nhập/xuất của tài khoản (cũ -> mới). STT 3 chữ số (-001) tự đổi sang 2 chữ số (-01)."""
    ten = _ten_khoa()
    with _db() as con:
        rows = con.execute("SELECT thoi_gian, hanh_dong, ma_san_pham, ten_hang, loai FROM lich_su "
                           "WHERE tai_khoan = ? ORDER BY id", (ten,)).fetchall()
    lich_su = [dict(r) for r in rows]
    for muc in lich_su:
        if muc.get("ma_san_pham"):
            muc["ma_san_pham"] = _dinh_dang_stt_trong_ma(muc["ma_san_pham"])
    return lich_su


def ghi_lich_su(hanh_dong, ma_sp, ten_hang, loai):
    ten = _ten_khoa()
    with _db() as con:
        con.execute("INSERT INTO lich_su(tai_khoan, thoi_gian, hanh_dong, ma_san_pham, ten_hang, loai) VALUES (?,?,?,?,?,?)",
                    (ten, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), hanh_dong, ma_sp, ten_hang, loai))


# ---------- Tên sản phẩm của sản phẩm (mã sản phẩm trong mã QR -> tên sản phẩm để hiển thị) ----------
def luu_ten_day_du(cac_cap):
    """Lưu / cập nhật danh sách (mã sản phẩm, tên sản phẩm) của tài khoản hiện tại."""
    cac_cap = [(chuan_hoa(vt), dd.strip()) for vt, dd in cac_cap if chuan_hoa(vt) and dd and dd.strip()]
    if not cac_cap:
        return
    ten = _ten_khoa()
    with _db() as con:
        con.executescript(SCHEMA_TEN_SAN_PHAM)                 # bảo đảm bảng tồn tại (CSDL tạo từ bản cũ)
        con.executemany(
            "INSERT INTO ten_san_pham(tai_khoan, viet_tat, day_du) VALUES (?,?,?) "
            "ON CONFLICT(tai_khoan, viet_tat) DO UPDATE SET day_du = excluded.day_du",
            [(ten, vt, dd) for vt, dd in cac_cap])


def doc_ten_day_du():
    """Trả về {mã sản phẩm đã chuẩn hóa: tên sản phẩm} của tài khoản hiện tại."""
    with _db() as con:
        con.executescript(SCHEMA_TEN_SAN_PHAM)
        return {r["viet_tat"]: r["day_du"]
                for r in con.execute("SELECT viet_tat, day_du FROM ten_san_pham WHERE tai_khoan = ?", (_ten_khoa(),))}


def ten_day_du(ten_hang, bang=None):
    """Tên hiển thị của sản phẩm: có tên sản phẩm thì dùng, không thì giữ nguyên mã sản phẩm đang lưu.
    `bang` = kết quả doc_ten_day_du() (truyền vào khi gọi trong vòng lặp để khỏi đọc DB nhiều lần)."""
    if bang is None:
        bang = doc_ten_day_du()
    return bang.get(chuan_hoa(ten_hang), ten_hang)


def luu_ten_loai_day_du(cac_bo):
    """Lưu / cập nhật danh sách (mã sản phẩm, mã loại, tên loại) của tài khoản hiện tại."""
    cac_bo = [(chuan_hoa(sp), chuan_hoa(lo), tl.strip()) for sp, lo, tl in cac_bo
              if chuan_hoa(sp) and chuan_hoa(lo) and tl and tl.strip()]
    if not cac_bo:
        return
    ten = _ten_khoa()
    with _db() as con:
        con.executescript(SCHEMA_TEN_SAN_PHAM)
        con.executemany(
            "INSERT INTO loai_san_pham(tai_khoan, ma_sp, ma_loai, ten_loai) VALUES (?,?,?,?) "
            "ON CONFLICT(tai_khoan, ma_sp, ma_loai) DO UPDATE SET ten_loai = excluded.ten_loai",
            [(ten, sp, lo, tl) for sp, lo, tl in cac_bo])


def doc_ten_loai_day_du():
    """Trả về {(mã sản phẩm, mã loại) đã chuẩn hóa: tên loại} của tài khoản hiện tại."""
    with _db() as con:
        con.executescript(SCHEMA_TEN_SAN_PHAM)
        return {(r["ma_sp"], r["ma_loai"]): r["ten_loai"]
                for r in con.execute("SELECT ma_sp, ma_loai, ten_loai FROM loai_san_pham WHERE tai_khoan = ?", (_ten_khoa(),))}


def ten_loai_day_du(ten_hang, loai, bang=None):
    """Tên loại hiển thị: có tên loại thì dùng, không thì giữ nguyên mã loại đang lưu.
    `bang` = kết quả doc_ten_loai_day_du() (truyền vào khi gọi trong vòng lặp)."""
    if bang is None:
        bang = doc_ten_loai_day_du()
    return bang.get((chuan_hoa(ten_hang), chuan_hoa(loai)), loai)


# ============ GHI NHỚ ĐĂNG NHẬP (sống sót khi tải lại trang) ============
# st.session_state chỉ sống trong 1 phiên websocket: F5 / tải lại trang = phiên mới = state rỗng = bị văng ra.
# Cách sửa: sau khi đăng nhập, cấp 1 mã phiên ngẫu nhiên -> lưu BẢN BĂM trong bảng phien + lưu mã gốc vào cookie trình duyệt.
# Mỗi lần mở / tải lại trang, đọc cookie, đối chiếu với bảng phien, hợp lệ thì tự đăng nhập lại.
def _bam_phien(ma):
    return hashlib.sha256(ma.encode("utf-8")).hexdigest()


def tao_phien(ten):
    """Tạo mã phiên mới cho tài khoản `ten`, trả về mã gốc (để đặt vào cookie)."""
    now = time.time()
    ma = secrets.token_urlsafe(32)
    with _db() as con:
        con.execute("DELETE FROM phien WHERE het_han <= ?", (now,))              # dọn phiên hết hạn
        con.execute("INSERT INTO phien(ma_bam, ten, het_han) VALUES (?,?,?)",
                    (_bam_phien(ma), ten, now + PHIEN_NGAY * 86400))
    return ma


def tim_phien(ma):
    """Mã phiên hợp lệ + tài khoản còn tồn tại -> trả về tên tài khoản, ngược lại None."""
    if not ma:
        return None
    with _db() as con:
        r = con.execute("SELECT ten FROM phien WHERE ma_bam = ? AND het_han > ?", (_bam_phien(ma), time.time())).fetchone()
    return r["ten"] if r else None


def huy_phien(ma):
    if not ma:
        return
    with _db() as con:
        con.execute("DELETE FROM phien WHERE ma_bam = ?", (_bam_phien(ma),))


def _cookie_hien_tai():
    try:
        ma = st.context.cookies.get(COOKIE_PHIEN)      # cần Streamlit >= 1.37
        return ma if isinstance(ma, str) and ma else None
    except Exception:
        return None


def _js_cookie(ma=None):
    """Chạy 1 đoạn JS nhỏ để đặt (ma) hoặc xóa (None) cookie phiên trên trình duyệt."""
    if ma:
        lenh = f"{COOKIE_PHIEN}={ma}; path=/; max-age={PHIEN_NGAY * 86400}; SameSite=Lax"
    else:
        lenh = f"{COOKIE_PHIEN}=; path=/; max-age=0; SameSite=Lax"
    components.html(f"<script>window.parent.document.cookie = {json.dumps(lenh)};</script>", height=0)


# ============ KHỞI TẠO SESSION STATE ============
if st.session_state.get("nguoi_dung") and "kho" not in st.session_state:
    try:
        st.session_state.kho = doc_du_lieu()
    except PermissionError:           # tài khoản không còn tồn tại -> coi như chưa đăng nhập
        st.session_state.clear()

if "chuc_nang" not in st.session_state:
    st.session_state.chuc_nang = None


def co_ke():
    """Đã có ít nhất 1 kệ hàng chưa? (chưa có thì không cho Nhập / Xuất)"""
    return len(st.session_state.kho) > 0


def ke_dang_dung(ten_ke):
    """Số vị trí đang có hàng trong 1 kệ."""
    return sum(1 for hang in st.session_state.kho[ten_ke] for o in hang for vt in o if vt is not None)


def xoa_ke(ten_ke):
    """Xóa 1 kệ (chỉ khi kệ trống), lưu file, dọn các lựa chọn đang trỏ tới kệ đó. Trả về True nếu xóa được."""
    if ten_ke not in st.session_state.kho or ke_dang_dung(ten_ke) > 0:
        return False
    del st.session_state.kho[ten_ke]
    luu_du_lieu(st.session_state.kho)
    for k in ("nhap_ke", "xuat_ke", "kt_ke"):               # chọn ô thủ công đang ở kệ vừa xóa
        if st.session_state.get(k) == ten_ke:
            for kk in (k, k.replace("_ke", "_hang"), k.replace("_ke", "_cot")):
                st.session_state.pop(kk, None)
    for k in ("nhap_o", "kt_qr_o"):                         # ô đang chọn bằng quét QR
        o = st.session_state.get(k)
        if o is not None and o[0] == ten_ke:
            st.session_state.pop(k, None)
    return True


def so_ke_tiep_theo():
    return max((lay_so_ke(k) for k in st.session_state.kho), default=0) + 1


def them_ke(so_hang, so_cot, so_vi_tri=SO_VI_TRI):
    """Tạo kệ mới (Kệ n+1) với số hàng x số cột x số vị trí mỗi ô đã nhập, lưu file, trả về tên kệ."""
    ten_ke = f"Kệ {so_ke_tiep_theo()}"
    st.session_state.kho[ten_ke] = tao_ke_rong(int(so_hang), int(so_cot), int(so_vi_tri))
    luu_du_lieu(st.session_state.kho)
    return ten_ke


def kiem_tra_sua_ke(ten_ke, so_hang, so_cot, so_vt):
    """Trả về danh sách lý do KHÔNG thể đổi kệ sang kích thước mới (rỗng = sửa được).
    Chỉ cắt bớt tầng / cột / vị trí ở phía cuối, và chỉ khi phần bị cắt đang trống."""
    ke = st.session_state.kho[ten_ke]
    cu_hang, cu_cot, cu_vt = len(ke), len(ke[0]), len(ke[0][0])
    vuong = {"hang": [], "cot": [], "vt": []}
    for h in range(cu_hang):
        for c in range(cu_cot):
            o = ke[h][c]
            if h >= so_hang:
                if any(vt is not None for vt in o):
                    vuong["hang"].append(ma_o(ten_ke, h, c))
            elif c >= so_cot:
                if any(vt is not None for vt in o):
                    vuong["cot"].append(ma_o(ten_ke, h, c))
            elif so_vt < cu_vt and any(vt is not None for vt in o[so_vt:]):
                vuong["vt"].append(ma_o(ten_ke, h, c))

    def liet_ke(ds):
        return ", ".join(ds[:8]) + (f" và {len(ds) - 8} ô khác" if len(ds) > 8 else "")

    ly_do = []
    if vuong["hang"]:
        ly_do.append(f"Các tầng từ {so_hang + 1} trở lên còn hàng ở: {liet_ke(vuong['hang'])}.")
    if vuong["cot"]:
        ly_do.append(f"Các cột từ {so_cot + 1} trở lên còn hàng ở: {liet_ke(vuong['cot'])}.")
    if vuong["vt"]:
        ly_do.append(f"Vị trí từ {so_vt + 1} trở lên còn hàng ở: {liet_ke(vuong['vt'])}.")
    return ly_do


def sua_ke(ten_ke, so_hang, so_cot, so_vt):
    """Đổi kích thước 1 kệ (thêm / bớt tầng, cột, vị trí mỗi ô). Hàng đang có giữ nguyên vị trí và mã.
    Trả về (True, "") hoặc (False, lý do)."""
    if ten_ke not in st.session_state.kho:
        return False, "Kệ không tồn tại."
    so_hang, so_cot, so_vt = int(so_hang), int(so_cot), int(so_vt)
    ly_do = kiem_tra_sua_ke(ten_ke, so_hang, so_cot, so_vt)
    if ly_do:
        return False, " ".join(ly_do)
    cu = st.session_state.kho[ten_ke]
    moi = tao_ke_rong(so_hang, so_cot, so_vt)
    for h in range(min(so_hang, len(cu))):
        for c in range(min(so_cot, len(cu[0]))):
            for i in range(min(so_vt, len(cu[h][c]))):
                moi[h][c][i] = cu[h][c][i]
    st.session_state.kho[ten_ke] = moi
    luu_du_lieu(st.session_state.kho)
    for k in ("nhap_ke", "xuat_ke", "kt_ke"):                 # bỏ ô đang chọn thủ công ở kệ này (kích thước đã đổi)
        if st.session_state.get(k) == ten_ke:
            for kk in (k.replace("_ke", "_hang"), k.replace("_ke", "_cot")):
                st.session_state.pop(kk, None)
    return True, ""


# ============ SINH MÃ Ô & MÃ SẢN PHẨM ============
def lay_so_ke(ten_ke):
    return int(ten_ke.replace("Kệ", "").strip())


def ma_o(ten_ke, h, c):
    """Mã ô: K{kệ}T{tầng}{ô}. Ví dụ: K1T101"""
    so_ke = lay_so_ke(ten_ke)
    return f"K{so_ke}T{h + 1}{c + 1:02d}"


def ma_san_pham_tu_so_ke(so_ke, ten_hang, loai, stt):
    """Mã sản phẩm: K{kệ}-{TÊN SP}-{LOẠI}-{STT}. Ví dụ: K1-VFD-3P-01"""
    return f"K{so_ke}-{chuan_hoa(ten_hang)}-{chuan_hoa(loai)}-{stt:02d}"


def ma_san_pham(ten_ke, h, c, ten_hang, loai, stt):
    """Mã sản phẩm đầy đủ: K{kệ}-{TÊN SP}-{LOẠI}-{STT} (VD: K1-VFD-3P-01). h, c giữ lại cho tương thích, không dùng."""
    return ma_san_pham_tu_so_ke(lay_so_ke(ten_ke), ten_hang, loai, stt)


# ============ GIAO DIỆN ============
# Ý tưởng: kho thật là kệ sắt — cột xanh, thanh ngang cam, mỗi ô chia thành các vị trí.
# Giao diện vẽ đúng như vậy: mỗi ô có một dãy vạch = số vị trí, vạch xanh là đã có hàng.
# Phẳng, không đổ bóng, không gradient, không icon emoji. Tự đổi sang tông tối theo hệ điều hành.

CSS_CHUNG = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Barlow+Semi+Condensed:wght@400;500;600;700&display=swap');
:root{
  --nen:#F7F5F0;   /* màu nền trang: be nhạt */
  --bg:#e9ebee; --paper:#f8f9fa; --ink:#14181f; --muted:#5a6270; --line:#c5cbd4; --slot:#d3d8df;
  --rack:#1a3c9c; --on-rack:#ffffff; --rack-dark:#12296e; --rack-tint:#dfe6f7; --beam:#ee7a14;
  --bad:#b3261e; --bad-tint:#f6dcd9;
  --tc:#0f766e; --tc-dark:#0a5750; --on-tc:#ffffff;      /* Tra cứu: xanh ngọc */
  --tl:#6d3fb8; --tl-dark:#522e8c; --on-tl:#ffffff;      /* Thiết lập: tím */
  --f-ui:'Barlow Semi Condensed','Segoe UI',Roboto,'Helvetica Neue',sans-serif;
  --f-head:'Barlow Condensed','Arial Narrow','Segoe UI',sans-serif;
}
@media (prefers-color-scheme: dark){
  :root{
    --nen:#0f1218;
    --bg:#0f1218; --paper:#171b23; --ink:#e7eaef; --muted:#97a0ad; --line:#2f3643; --slot:#2a303b;
    --rack:#86a2ff; --on-rack:#0b1226; --rack-dark:#a9bdff; --rack-tint:#1b2440; --beam:#f28b2d;
    --bad:#ff8a80; --bad-tint:#3b1d1b;
    --tc:#5fd6c8; --tc-dark:#8fe6dc; --on-tc:#06201d;
    --tl:#c2a6f5; --tl-dark:#d9c8fa; --on-tl:#1a0f33;
  }
}
.stApp, [data-testid="stAppViewContainer"]{ background:var(--nen) !important; }
.stApp{ color:var(--ink); }
.stApp, .stApp p, .stApp label, .stApp li, .stApp td, .stApp th, .stApp input, .stApp textarea, .stApp button, .stApp select,
.stApp [data-baseweb], .stApp [data-testid="stCaptionContainer"]{ font-family:var(--f-ui); }
.stApp span[data-testid="stIconMaterial"]{ font-family:"Material Symbols Rounded" !important; }
[data-testid="stMarkdownContainer"] p{ font-size:1.05rem; }
[data-testid="stHeader"]{ background:transparent !important; }
[data-testid="stMainBlockContainer"], .block-container{ max-width:1180px; margin:0 auto; padding-top:2.4rem; }
div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdownContainer"] > style),
div.element-container:has(> div[data-testid="stMarkdownContainer"] > style){ display:none; }
h1,h2,h3,h4,h5{ font-family:var(--f-head); font-weight:600; letter-spacing:0; color:var(--ink); }
h5{ font-size:1.35rem; }
[data-testid="stWidgetLabel"] p, label p{ color:var(--ink); font-weight:500; }
[data-testid="stCaptionContainer"]{ color:var(--muted); }
[data-testid="stMarkdownContainer"] hr, hr{ border:0; border-top:1px solid var(--line); }
:focus-visible{ outline:2px solid var(--rack); outline-offset:2px; }

/* ---- tiêu đề trang ---- */
.pg-head{ display:flex; align-items:baseline; justify-content:space-between; gap:16px; flex-wrap:wrap;
  border-bottom:2px solid var(--ink); padding-bottom:10px; margin:0 0 20px; }
.pg-title{ font:600 clamp(1.7rem,5vw,2.6rem)/1.1 var(--f-head); color:var(--ink); }
.pg-meta{ color:var(--muted); font-size:1rem; }
.pg-head.nho{ padding-bottom:6px; margin-bottom:12px; border-bottom-width:1px; }
.pg-head.nho .pg-title{ font-size:clamp(1.3rem,4vw,1.7rem); }
.sec{ margin:22px 0 8px; font:600 1.2rem/1.2 var(--f-head); color:var(--ink); }

/* ---- dải số liệu (một khối liền, ngăn bằng đường kẻ) ---- */
.strip{ display:grid; grid-template-columns:repeat(4,1fr); border:1px solid var(--line); background:var(--paper);
  border-radius:8px; margin:0 0 22px; }
.strip > div{ padding:14px 18px 16px; border-left:1px solid var(--line); min-width:0; }
.strip > div:first-child{ border-left:0; }
.strip .v{ display:block; font:600 2.1rem/1 var(--f-head); font-variant-numeric:tabular-nums; color:var(--ink); }
.strip .v small{ font-size:1.05rem; font-weight:500; color:var(--muted); }
.strip .k{ display:block; margin-top:7px; color:var(--ink); font-size:.95rem; font-weight:500; }
.ico{ width:18px; height:18px; flex:none; vertical-align:-3px; margin-right:8px; }
.ico.c-blue{ color:var(--rack); } .ico.c-green{ color:var(--tc); }
.ico.c-orange{ color:var(--beam); } .ico.c-violet{ color:var(--tl); }
.pg-title .ico-img{ height:30px; width:auto; vertical-align:-6px; margin-right:10px; }
.pg-title .ico{ width:26px; height:26px; vertical-align:-4px; margin-right:10px; }
.strip .s{ display:block; margin-top:1px; color:var(--muted); font-size:.85rem; }
.strip .bar{ height:4px; background:var(--slot); margin-top:10px; border-radius:2px; overflow:hidden; }
.strip .bar i{ display:block; height:100%; background:var(--rack); }
.canh-bao{ border-left:4px solid var(--bad); background:var(--bad-tint); color:var(--ink); padding:10px 14px;
  border-radius:0 6px 6px 0; margin:-8px 0 20px; font-weight:500; }
@media (max-width:700px){
  .strip{ grid-template-columns:repeat(2,1fr); }
  .strip > div:nth-child(odd){ border-left:0; }
  .strip > div:nth-child(n+3){ border-top:1px solid var(--line); }
}

/* ---- nút ---- */
div[data-testid="stButton"] > button, div[data-testid="stDownloadButton"] > button, div[data-testid="stFormSubmitButton"] > button{
  min-height:44px; border-radius:6px; font-weight:600; background-color:var(--paper); color:var(--ink);
  border:1px solid var(--line); box-shadow:none; transition:background-color .12s, border-color .12s;
}
div[data-testid="stButton"] > button p, div[data-testid="stDownloadButton"] > button p, div[data-testid="stFormSubmitButton"] > button p{
  color:inherit; font-weight:600; font-size:1.04rem; }
div[data-testid="stButton"] > button:hover, div[data-testid="stDownloadButton"] > button:hover, div[data-testid="stFormSubmitButton"] > button:hover{
  border-color:var(--ink); background-color:var(--bg); color:var(--ink); }
div[data-testid="stButton"] > button:disabled{ opacity:.45; }
div[data-testid="stButton"] > button[kind^="primary"], div[data-testid="stButton"] > button[data-testid^="stBaseButton-primary"],
div[data-testid="stDownloadButton"] > button[kind^="primary"], div[data-testid="stDownloadButton"] > button[data-testid^="stBaseButton-primary"],
div[data-testid="stFormSubmitButton"] > button[kind^="primary"], div[data-testid="stFormSubmitButton"] > button[data-testid^="stBaseButton-primary"]{
  background-color:var(--rack); color:var(--on-rack); border-color:var(--rack); }
div[data-testid="stButton"] > button[kind^="primary"]:hover, div[data-testid="stButton"] > button[data-testid^="stBaseButton-primary"]:hover,
div[data-testid="stDownloadButton"] > button[kind^="primary"]:hover, div[data-testid="stDownloadButton"] > button[data-testid^="stBaseButton-primary"]:hover,
div[data-testid="stFormSubmitButton"] > button[kind^="primary"]:hover, div[data-testid="stFormSubmitButton"] > button[data-testid^="stBaseButton-primary"]:hover{
  background-color:var(--rack-dark); color:var(--on-rack); border-color:var(--rack-dark); }
.st-key-dang_xuat button{ min-height:36px !important; }

/* ---- ô nhập, tab, thông báo, form, bảng ---- */
div[data-baseweb="input"], div[data-baseweb="base-input"], div[data-baseweb="textarea"], div[data-baseweb="select"] > div{
  border-radius:6px !important; }
div[data-baseweb="input"], div[data-baseweb="textarea"], div[data-baseweb="select"] > div{
  background-color:var(--paper) !important; border-color:var(--line) !important; }
div[data-baseweb="tab-list"]{ gap:4px; border-bottom:1px solid var(--line); }
div[data-baseweb="tab-border"]{ background:transparent !important; }
button[data-baseweb="tab"]{ border-radius:0; font-weight:600; padding:8px 14px; color:var(--muted); }
button[data-baseweb="tab"][aria-selected="true"]{ color:var(--ink); }
div[data-baseweb="tab-highlight"]{ background-color:var(--rack) !important; height:3px; }
div[data-testid="stAlert"]{ border-radius:6px; box-shadow:none; border:1px solid var(--line); }
.st-key-xoa_ke_loi div[data-testid="stAlert"]{ background:var(--bad-tint); border-color:var(--bad); }
.st-key-xoa_ke_loi div[data-testid="stAlert"] *{ color:var(--bad) !important; }
div[data-testid="stForm"]{ background:var(--paper); border:1px solid var(--line); border-radius:8px; padding:18px; }
div[data-testid="stExpander"] details{ border-radius:8px; border:1px solid var(--line); background:var(--paper); }
div[data-testid="stDataFrame"]{ border-radius:8px; overflow:hidden; border:1px solid var(--line); }

/* ---- sơ đồ kệ chọn ô (nút bấm) ---- */
div[class*="st-key-sodo_"] div[data-testid="stHorizontalBlock"]{ gap:6px; flex-wrap:nowrap !important; }
div[class*="st-key-sodo_"] div[data-testid="stColumn"]{ min-width:0 !important; flex:1 1 0 !important; width:auto !important; }
div[class*="st-key-sodo_"] div[data-testid="stColumn"]:first-child{ flex:0 0 30px !important; }
div[class*="st-key-sodo_"] div[data-testid="stButton"] > button{
  min-height:62px; padding:6px 2px 14px; border-radius:3px; line-height:1.2;
  background-repeat:no-repeat; background-size:100% 6px; background-position:left bottom; }
div[class*="st-key-sodo_"] div[data-testid="stButton"] > button p{ font-size:.82rem; line-height:1.3; text-align:center; }
.sodo-t{ display:flex; align-items:center; height:62px; font:600 .95rem var(--f-head); color:var(--muted); }

@media (max-width:640px){
  [data-testid="stMainBlockContainer"], .block-container{ padding-top:1.4rem; }
  .strip > div{ padding:12px 12px 14px; } .strip .v{ font-size:1.8rem; }
}
</style>
"""

# CSS RIÊNG của trang chủ: cột thao tác + kệ vẽ bằng HTML.
CSS_TRANG_CHU = """
<style>
.st-key-home_actions button{ min-height:56px; padding:0 18px; justify-content:flex-start; }
.st-key-home_actions button > div{ width:100%; justify-content:flex-start; }
.st-key-home_actions button div[data-testid="stMarkdownContainer"]{ width:100%; }
.st-key-home_actions button p{ text-align:left; font-size:1.15rem; }
.st-key-home_actions .sec:first-child{ margin-top:0; }
.st-key-home_actions .st-key-home_kiem_tra div[data-testid="stButton"] > button,
.st-key-home_actions .st-key-home_lich_su div[data-testid="stButton"] > button{
  background-color:var(--tc); color:var(--on-tc); border-color:var(--tc); }
.st-key-home_actions .st-key-home_kiem_tra div[data-testid="stButton"] > button:hover,
.st-key-home_actions .st-key-home_lich_su div[data-testid="stButton"] > button:hover{
  background-color:var(--tc-dark); color:var(--on-tc); border-color:var(--tc-dark); }
.st-key-home_actions .st-key-home_tao_ke div[data-testid="stButton"] > button,
.st-key-home_actions .st-key-home_tao_qr div[data-testid="stButton"] > button{
  background-color:var(--tl); color:var(--on-tl); border-color:var(--tl); }
.st-key-home_actions .st-key-home_tao_ke div[data-testid="stButton"] > button:hover,
.st-key-home_actions .st-key-home_tao_qr div[data-testid="stButton"] > button:hover{
  background-color:var(--tl-dark); color:var(--on-tl); border-color:var(--tl-dark); }
.racks{ display:flex; flex-wrap:wrap; gap:26px 24px; }
.rack{ flex:1 1 320px; min-width:0; }
div[class*="st-key-sodo_home_"]{ background:var(--paper); border-left:7px solid var(--rack); border-right:7px solid var(--rack);
  border-top:4px solid var(--beam); border-bottom:4px solid var(--beam); padding:8px 6px 6px; margin-bottom:22px; }
.rack-top{ display:flex; justify-content:space-between; align-items:baseline; gap:10px; flex-wrap:wrap; margin-bottom:6px; }
.rack-top b{ font:600 1.35rem/1 var(--f-head); color:var(--ink); }
.rack-top span{ color:var(--muted); font-size:.92rem; }
.frame{ background:var(--paper); border-left:7px solid var(--rack); border-right:7px solid var(--rack);
  border-top:4px solid var(--beam); overflow-x:auto; }
.tier{ display:grid; grid-template-columns:30px 1fr; align-items:center; border-bottom:4px solid var(--beam); padding:6px 4px 6px 0; }
.tier .tl{ font:600 .9rem var(--f-head); color:var(--muted); text-align:center; }
.cells{ display:grid; gap:4px; }
.cell{ background:var(--bg); border-radius:2px; padding:5px 5px 6px; min-width:0; }
.cell .cc{ display:block; font-size:.72rem; font-weight:600; color:var(--muted); font-variant-numeric:tabular-nums; white-space:nowrap; }
.cell .ticks{ display:grid; grid-auto-flow:column; grid-auto-columns:1fr; gap:2px; margin-top:4px; }
.cell .ticks i{ display:block; height:7px; background:var(--slot); }
.cell .ticks i.on{ background:var(--rack); }
.cell.full{ background:var(--rack-tint); }
.cell.full .cc{ color:var(--ink); }
.chu-thich{ color:var(--muted); font-size:.9rem; margin:10px 0 0; }
</style>
"""


def _gon_html(html):
    """Bỏ thụt đầu dòng/xuống dòng để Markdown không hiểu nhầm đoạn HTML là khối code."""
    return "".join(dong.strip() for dong in html.splitlines())


_EMOJI = re.compile("[\U0001F300-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\uFE0F\u200d]")


def _dep(chuoi):
    """Bỏ emoji; chữ IN HOA toàn bộ thì đổi sang kiểu viết câu (giữ nguyên 'QR')."""
    s = re.sub(r"\s+", " ", _EMOJI.sub("", chuoi)).strip()
    if s and s.isupper():
        s = re.sub(r"\bQr\b", "QR", s.capitalize())
    return s


def tieu_de_trang(tieu_de, nho=False, phu=""):
    """Tiêu đề trang: chữ lớn, một đường kẻ đậm bên dưới. phu = dòng nhỏ bên phải (tùy chọn). nho=True: bản gọn."""
    meta = f'<span class="pg-meta">{_esc(phu)}</span>' if phu else ""
    st.markdown(_gon_html(
        f'<div class="pg-head{" nho" if nho else ""}"><div class="pg-title">{_esc(_dep(tieu_de))}</div>{meta}</div>'),
        unsafe_allow_html=True)


def nhan_muc(ten):
    """Tên nhóm nội dung, ví dụ: Bộ lọc, Kết quả."""
    st.markdown(f'<div class="sec">{_esc(ten)}</div>', unsafe_allow_html=True)


# ============ ICON ============
# Icon nét mảnh (SVG nội tuyến, tô màu theo chữ) cho dải số liệu và tiêu đề; icon nút dùng bộ Material của Streamlit.
_ICON = {
    "box": '<path d="M11 21.73a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73z"/><path d="M12 22V12"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="m7.5 4.27 9 5.15"/>',
    "pin": '<path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0Z"/><circle cx="12" cy="10" r="3"/>',
    "square": '<rect width="18" height="18" x="3" y="3" rx="2"/>',
    "lock": '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
    "calendar": '<path d="M8 2v4"/><path d="M16 2v4"/><rect width="18" height="18" x="3" y="4" rx="2"/><path d="M3 10h18"/><path d="M8 14h.01"/><path d="M12 14h.01"/><path d="M16 14h.01"/><path d="M8 18h.01"/><path d="M12 18h.01"/><path d="M16 18h.01"/>',
}


def svg_icon(ten, mau="blue"):
    """Icon SVG nội tuyến; `mau`: blue / green / orange / violet (xem CSS .ico.c-...)."""
    d = _ICON.get(ten)
    if not d:
        return ""
    return (f'<svg class="ico c-{mau}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
            f'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{d}</svg>')


def icon_kw(ten):
    """Tham số icon cho st.button (Material). Streamlit < 1.37 chưa hỗ trợ -> bỏ qua, nút vẫn chạy bình thường."""
    try:
        if tuple(int(x) for x in st.__version__.split(".")[:2]) >= (1, 37):
            return {"icon": f":material/{ten}:"}
    except Exception:
        pass
    return {}


# Icon cạnh tiêu đề "Tổng quan kho" ở trang chủ:
ICON_TONG_QUAN = ""         # icon có sẵn (box / pin / square / lock / calendar); để "" nếu không muốn icon
LOGO_TONG_QUAN = ""         # muốn dùng ảnh riêng: điền đường dẫn ảnh (png / jpg / svg), VD "logo.png" (đặt cạnh file app); ưu tiên hơn dòng trên


def icon_tieu_de():
    """HTML icon đặt trước tiêu đề trang chủ: ảnh riêng nếu có, không thì icon có sẵn."""
    if LOGO_TONG_QUAN and os.path.isfile(LOGO_TONG_QUAN):
        try:
            import base64
            mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "svg": "image/svg+xml",
                    "webp": "image/webp"}.get(LOGO_TONG_QUAN.rsplit(".", 1)[-1].lower(), "image/png")
            with open(LOGO_TONG_QUAN, "rb") as f:
                du_lieu = base64.b64encode(f.read()).decode("ascii")
            return f'<img class="ico-img" src="data:{mime};base64,{du_lieu}" alt="">'
        except OSError:
            pass
    return svg_icon(ICON_TONG_QUAN, "blue")


def html_the_so_lieu(cac_the):
    """Dải số liệu liền khối: mỗi phần tử (tên icon, số, nhãn, ghi chú, màu icon)."""
    o = "".join(
        f'<div><span class="v">{so}</span><span class="k">{svg_icon(icon, mau)}{nhan}</span><span class="s">{ghi_chu}</span></div>'
        for icon, so, nhan, ghi_chu, mau in cac_the)
    return _gon_html(f'<div class="strip" style="grid-template-columns:repeat({len(cac_the)},1fr)">{o}</div>')


def thong_ke_kho():
    """Số liệu thật của kho để hiển thị ở trang chủ."""
    tong_o = sum(len(hang) for ke in st.session_state.kho.values() for hang in ke)
    tong_vi_tri = sum(len(o) for ke in st.session_state.kho.values() for hang in ke for o in hang)
    da_dung = 0
    o_day = 0
    for ke in st.session_state.kho.values():
        for hang in ke:
            for o in hang:
                so = sum(1 for vt in o if vt is not None)
                da_dung += so
                if so == len(o):
                    o_day += 1
    hom_nay = datetime.now().strftime("%Y-%m-%d")
    nhap_hn = xuat_hn = 0
    for m in doc_lich_su():
        if str(m.get("thoi_gian", "")).startswith(hom_nay):
            if m.get("hanh_dong") == "Nhập":
                nhap_hn += 1
            elif m.get("hanh_dong") == "Xuất":
                xuat_hn += 1
    return {
        "tong_o": tong_o, "tong_vi_tri": tong_vi_tri, "da_dung": da_dung,
        "con_trong": tong_vi_tri - da_dung, "ti_le": round(da_dung * 100 / tong_vi_tri) if tong_vi_tri else 0,
        "o_day": o_day, "nhap_hn": nhap_hn, "xuat_hn": xuat_hn,
        "hom_nay": datetime.now().strftime("%d/%m/%Y"),
    }


_THU = ["Thứ hai", "Thứ ba", "Thứ tư", "Thứ năm", "Thứ sáu", "Thứ bảy", "Chủ nhật"]


def html_dau_trang_chu(tk):
    """Tiêu đề trang chủ + ngày hôm nay."""
    ngay = f"{_THU[datetime.now().weekday()]}, {tk['hom_nay']}"
    return _gon_html(f'<div class="pg-head"><div class="pg-title">{icon_tieu_de()}Tổng quan kho</div><span class="pg-meta">{ngay}</span></div>')


def html_dai_so_lieu(tk):
    """Dải 4 số liệu chính của kho + cảnh báo khi gần đầy."""
    canh_bao = ""
    if tk["ti_le"] >= 90:
        canh_bao = f'<div class="canh-bao">Kho đã dùng {tk["ti_le"]}% vị trí. Cần xuất bớt hàng hoặc tạo thêm kệ.</div>'
    rong = max(tk["ti_le"], 2) if tk["da_dung"] > 0 else 0       # có hàng thì luôn thấy một đoạn thanh
    return _gon_html(f"""
    <div class="strip">
      <div><span class="v">{tk['da_dung']}<small> / {tk['tong_vi_tri']}</small></span>
        <span class="k">{svg_icon("pin", "blue")}Vị trí đang chứa hàng</span>
        <div class="bar"><i style="width:{rong}%"></i></div></div>
      <div><span class="v">{tk['con_trong']}</span><span class="k">{svg_icon("square", "green")}Vị trí còn trống</span>
        <span class="s">trên {tk['tong_o']} ô</span></div>
      <div><span class="v">{tk['o_day']}</span><span class="k">{svg_icon("lock", "orange")}Ô đã đầy</span></div>
      <div><span class="v">{tk['nhap_hn']} <small>nhập</small>  {tk['xuat_hn']} <small>xuất</small></span>
        <span class="k">{svg_icon("calendar", "violet")}Hôm nay</span></div>
    </div>{canh_bao}
    """)


def hien_o_trong_kho(ten_ke, h, c):
    """Bấm vào 1 ô ở trang chủ: mở hộp thoại liệt kê các sản phẩm đang nằm trong ô đó."""
    o = st.session_state.kho[ten_ke][h][c]
    tieu_de = f"Ô {ma_o(ten_ke, h, c)} — {sum(1 for vt in o if vt is not None)}/{len(o)} vị trí"

    def _noi_dung():
        hien_thi_danh_sach_ma_sp(ten_ke, h, c, co_tieu_de=False)

    if hasattr(st, "dialog"):
        try:
            st.dialog(tieu_de, width="large")(_noi_dung)()
        except TypeError:                 # bản Streamlit cũ chưa có tham số width
            st.dialog(tieu_de)(_noi_dung)()
    else:                                 # bản rất cũ chưa có st.dialog: hiện ngay trên trang
        st.markdown(f"##### {tieu_de}")
        _noi_dung()


def _vao_kiem_tra():
    st.session_state.chuc_nang = "kiem_tra"
    for k in ("kt_qr_tb", "kt_qr_o", "kt_qr_cuoi"):
        st.session_state.pop(k, None)
    st.session_state["kt_cam_n"] = st.session_state.get("kt_cam_n", 0) + 1


def _vao_nhap():
    st.session_state.chuc_nang = "nhap"
    for k in ("nhap_o", "nhap_tb", "nhap_qr_cuoi",      # bắt đầu lại từ bước quét QR ô
              "nhap_hang", "nhap_cot", "nhap_ke", "nhap_tay_tb"):
        st.session_state.pop(k, None)
    st.session_state["nhap_cam_n"] = st.session_state.get("nhap_cam_n", 0) + 1


def _vao_xuat():
    st.session_state.chuc_nang = "xuat"
    for k in ("xuat_tb", "xuat_ds", "xuat_cho_chon", "xuat_qr_cuoi",    # bắt đầu phiên xuất mới
              "xuat_hang", "xuat_cot", "xuat_ke", "xuat_tay_tb"):
        st.session_state.pop(k, None)
    st.session_state["xuat_cam_n"] = st.session_state.get("xuat_cam_n", 0) + 1


def _vao_lich_su():
    st.session_state.chuc_nang = "lich_su"


def _vao_tao_ke():
    st.session_state.chuc_nang = "tao_ke"
    st.session_state.pop("tao_ke_tb", None)


def _vao_tao_qr():
    st.session_state.chuc_nang = "tao_qr"
    st.session_state.pop("qr_ket_qua", None)
    st.session_state.pop("qro_ket_qua", None)


def trang_chu():
    st.markdown(_gon_html(CSS_TRANG_CHU), unsafe_allow_html=True)
    tk = thong_ke_kho()
    st.markdown(html_dau_trang_chu(tk), unsafe_allow_html=True)
    st.markdown(html_dai_so_lieu(tk), unsafe_allow_html=True)

    cot_tac_vu, cot_kho = st.columns([1, 2.6], gap="large")

    with cot_tac_vu:
        with st.container(key="home_actions"):
            nhan_muc("Hàng hóa")
            if st.button("Nhập hàng", use_container_width=True, key="home_nhap", **icon_kw("add_circle"), disabled=not co_ke(), type="primary"):
                _vao_nhap()
                st.rerun()
            if st.button("Xuất hàng", use_container_width=True, key="home_xuat", **icon_kw("remove_circle"), disabled=not co_ke(), type="primary"):
                _vao_xuat()
                st.rerun()

            nhan_muc("Tra cứu")
            if st.button("Kiểm tra kho", use_container_width=True, key="home_kiem_tra", **icon_kw("fact_check")):
                _vao_kiem_tra()
                st.rerun()
            if st.button("Lịch sử nhập xuất", use_container_width=True, key="home_lich_su", **icon_kw("history")):
                _vao_lich_su()
                st.rerun()

            nhan_muc("Thiết lập")
            if st.button("Kệ hàng", use_container_width=True, key="home_tao_ke", **icon_kw("shelves"),
                         type="secondary" if co_ke() else "primary"):
                _vao_tao_ke()
                st.rerun()
            if st.button("Tạo mã QR", use_container_width=True, key="home_tao_qr", **icon_kw("qr_code_2")):
                _vao_tao_qr()
                st.rerun()

    with cot_kho:
        if co_ke():
            o_chon = None
            for ten_ke in sorted(st.session_state.kho, key=lay_so_ke):
                o = ve_so_do_ke(ten_ke, key_prefix="home", kieu="home")
                if o and o_chon is None:
                    o_chon = (ten_ke, o[0], o[1])
            if o_chon:
                hien_o_trong_kho(*o_chon)


def khung_tinh_trang_kho(kieu):
    """Dải số liệu thật của kho, đặt dưới nút mở camera ở tab Quét QR của Nhập / Xuất / Kiểm tra."""
    tk = thong_ke_kho()
    if kieu == "nhap":
        cac_the = [
            ("square", tk["con_trong"], "Vị trí còn trống", f"trên tổng {tk['tong_vi_tri']} vị trí", "blue"),
            ("pin", tk["da_dung"], "Đang chứa hàng", "vị trí đã có sản phẩm", "green"),
            ("calendar", tk["nhap_hn"], "Nhập hôm nay", f"lượt, ngày {tk['hom_nay']}", "violet"),
        ]
    elif kieu == "xuat":
        cac_the = [
            ("pin", tk["da_dung"], "Đang chứa hàng", "sản phẩm có thể xuất", "blue"),
            ("square", tk["con_trong"], "Vị trí còn trống", f"trên tổng {tk['tong_vi_tri']} vị trí", "green"),
            ("calendar", tk["xuat_hn"], "Xuất hôm nay", f"lượt, ngày {tk['hom_nay']}", "violet"),
        ]
    else:
        cac_the = [
            ("box", tk["tong_o"], "Tổng số ô", f"{len(st.session_state.kho)} kệ hàng", "blue"),
            ("pin", tk["da_dung"], "Đang chứa hàng", "vị trí đã có sản phẩm", "green"),
            ("lock", tk["o_day"], "Ô đã đầy", "không còn vị trí trống", "orange"),
        ]
    nhan_muc("Tình trạng kho")
    st.markdown(html_the_so_lieu(cac_the), unsafe_allow_html=True)


def nut_quay_lai():
    """Nút quay về trang chủ — đặt ở CUỐI mỗi trang để dễ thao tác."""
    st.divider()
    if st.button("Quay lại trang chủ", use_container_width=True, key="quay_lai_trang_chu"):
        st.session_state.chuc_nang = None
        st.rerun()


# ============ VẼ SƠ ĐỒ KỆ HÀNG ============
def ve_so_do_ke(ten_ke, key_prefix="", kieu=""):
    """Sơ đồ chọn ô: mỗi ô là 1 nút; thanh mỏng dưới nút cho biết số vị trí đã dùng (CSS theo từng nút)."""
    du_lieu_ke = st.session_state.kho[ten_ke]
    o_duoc_chon = None

    so_hang, so_cot = len(du_lieu_ke), len(du_lieu_ke[0])
    so_ke = lay_so_ke(ten_ke)
    if kieu == "home":      # trang chủ: tên kệ bên trái, số vị trí đã dùng bên phải
        tong_ke = sum(len(o) for hang in du_lieu_ke for o in hang)
        da_ke = sum(1 for hang in du_lieu_ke for o in hang for vt in o if vt is not None)
        st.markdown(f'<div class="rack-top"><b>{_esc(ten_ke)}</b><span>đã dùng {da_ke}/{tong_ke}</span></div>',
                    unsafe_allow_html=True)
    else:
        st.markdown(f"**{ten_ke}**: {so_hang} hàng × {so_cot} cột")

    css_o = []
    with st.container(key=f"sodo_{key_prefix}_{so_ke}"):
        for h in range(so_hang):
            cols = st.columns([0.3] + [1] * so_cot)
            cols[0].markdown(f'<div class="sodo-t">T{h + 1}</div>', unsafe_allow_html=True)
            for c in range(so_cot):
                vi_tri_list = du_lieu_ke[h][c]
                nhan = ma_o(ten_ke, h, c)
                tong = len(vi_tri_list)
                so_dang_dung = sum(1 for vt in vi_tri_list if vt is not None)
                phan_tram = round(so_dang_dung * 100 / tong) if tong else 0
                khoa = f"{key_prefix}_k{so_ke}_{h}_{c}"
                sel = f'div.st-key-{khoa} div[data-testid="stButton"] > button'
                css_o.append(f"{sel}{{background-image:linear-gradient(90deg,var(--rack) {phan_tram}%,var(--slot) {phan_tram}%);}}")
                if tong and so_dang_dung == tong:
                    css_o.append(f"{sel}{{background-color:var(--rack-tint);border-color:var(--rack);}}")

                with cols[c + 1]:
                    if st.button(f"{nhan}  \n{so_dang_dung}/{tong}", key=khoa, use_container_width=True):
                        o_duoc_chon = (h, c)
    st.markdown(f"<style>{''.join(css_o)}</style>", unsafe_allow_html=True)
    return o_duoc_chon


def hien_thi_danh_sach_ma_sp(ten_ke, h, c, noi_bat=None, co_tieu_de=True):
    """Hiển thị đủ các vị trí của 1 ô — MỖI VỊ TRÍ CHỈ 1 MÃ SẢN PHẨM."""
    vi_tri_list = st.session_state.kho[ten_ke][h][c]
    if co_tieu_de:
        st.markdown(f"##### Danh sách sản phẩm trong ô **{ma_o(ten_ke, h, c)}**")

    bang_ten, bang_loai = doc_ten_day_du(), doc_ten_loai_day_du()
    for idx in range(len(vi_tri_list)):
        vt = vi_tri_list[idx]
        col_nhan, col_noi_dung = st.columns([1, 3])
        with col_nhan:
            st.write(f"**{'' if idx == noi_bat else ''}Vị trí {idx + 1}:**")
        with col_noi_dung:
            if vt is None:
                st.write("_(Trống - chưa có sản phẩm)_")
            else:
                ma = ma_san_pham(ten_ke, h, c, vt["ten_hang"], vt["loai"], vt["stt"])
                dong = f"`{ma}` — Tên: **{ten_day_du(vt['ten_hang'], bang_ten)}** — Loại: **{ten_loai_day_du(vt['ten_hang'], vt['loai'], bang_loai)}**"
                st.write(f"{dong}" if idx == noi_bat else dong)


# ============ CHỨC NĂNG: KIỂM TRA ============
# ============ KIỂM TRA: QUÉT QR VỊ TRÍ Ô ĐỂ XEM SẢN PHẨM TRONG Ô ============
def xu_ly_quet_kiem_tra(noi_dung):
    """Xử lý 1 lần quét ở chức năng kiểm tra: quét QR ô -> cho biết ô đó đang chứa những sản phẩm nào."""
    noi_dung = (noi_dung or "").strip()
    if not noi_dung:
        return

    o = phan_tich_ma_o(noi_dung)
    if o is None:
        st.session_state["kt_qr_tb"] = (
            "error", f"`{noi_dung}` không phải mã vị trí ô hợp lệ (VD: K1T101). "
                     f"Hãy quét QR **vị trí ô** để xem sản phẩm trong ô.")
        return

    ten_ke, h, c = o
    so_dung = sum(1 for vt in st.session_state.kho[ten_ke][h][c] if vt is not None)
    tong_vt = len(st.session_state.kho[ten_ke][h][c])
    st.session_state["kt_qr_o"] = o
    if so_dung == 0:
        st.session_state["kt_qr_tb"] = (
            "info", f"Ô **{ma_o(ten_ke, h, c)}** ({ten_ke}, tầng {h + 1}, ô số {c + 1})"
                    f"đang **trống** (0/{tong_vt} vị trí).")
    else:
        st.session_state["kt_qr_tb"] = (
            "success", f"Ô **{ma_o(ten_ke, h, c)}** ({ten_ke}, tầng {h + 1}, ô số {c + 1})"
                       f"đang có **{so_dung}/{tong_vt}** sản phẩm, còn {tong_vt - so_dung} vị trí trống.")


def _hien_thi_ket_qua_kiem_tra():
    tb = st.session_state.get("kt_qr_tb")
    if not tb:
        return
    {"success": st.success, "info": st.info, "warning": st.warning, "error": st.error}[tb[0]](tb[1])
    o = st.session_state.get("kt_qr_o")
    if o is not None:
        st.divider()
        hien_thi_danh_sach_ma_sp(*o)


def kiem_tra_bang_quet_qr():
    """Tab quét QR của trang kiểm tra (nút mở camera + kết quả lần quét gần nhất)."""
    if st.button("Mở camera quét QR", type="primary", use_container_width=True, key="kt_mo_cam"):
        st.session_state.pop("kt_qr_cuoi", None)
        st.session_state["kt_cam_n"] = st.session_state.get("kt_cam_n", 0) + 1  # camera mới
        st.session_state.chuc_nang = "quet_kiem_tra"
        st.rerun()
    khung_tinh_trang_kho("kiem_tra")
    if st.session_state.get("kt_qr_tb"):
        nhan_muc("Kết quả lần quét gần nhất")
    _hien_thi_ket_qua_kiem_tra()


def man_hinh_quet_kiem_tra():
    """Trang camera riêng cho kiểm tra ô."""
    if st.button("Quay lại trang kiểm tra", key="kt_quay_lai"):
        st.session_state.chuc_nang = "kiem_tra"
        st.rerun()
    tieu_de_trang("Quét QR kiểm tra ô", nho=True)

    if _QR_LIVE is None:
        st.error("Không tạo được thành phần camera (không ghi được file vào thư mục chạy app). "
                 "Hãy kiểm tra quyền ghi thư mục.")
        return

    # --- Camera quét liên tục (đặt trước các phần động để camera không bị tải lại) ---
    n = st.session_state.get("kt_cam_n", 0)
    ket_qua = _QR_LIVE(key=f"kt_scan_{n}", default=None)

    if ket_qua and ket_qua.get("id") != st.session_state.get("kt_qr_cuoi"):
        st.session_state["kt_qr_cuoi"] = ket_qua["id"]
        xu_ly_quet_kiem_tra(str(ket_qua.get("code", "")))

    st.subheader("Quét QR vị trí ô cần xem")
    _hien_thi_ket_qua_kiem_tra()
    st.divider()
    if st.button("Xong", use_container_width=True, key="kt_xong"):
        st.session_state.chuc_nang = "kiem_tra"
        st.rerun()


def man_hinh_kiem_tra():
    """Trang kiểm tra: chỉ quét QR vị trí ô để xem sản phẩm trong ô (xem sơ đồ kho đã chuyển ra trang chủ)."""
    tieu_de_trang("Kiểm tra kho hàng")
    kiem_tra_bang_quet_qr()
    nut_quay_lai()


# ============ PHÂN TÍCH NỘI DUNG QR ============
def phan_tich_ma_o(noi_dung):
    """Nếu nội dung là mã ô hợp lệ (VD: K1T101) -> trả về (tên kệ, tầng, ô), ngược lại None."""
    ma = noi_dung.strip().upper()
    for ten_ke, ke in st.session_state.kho.items():
        for h in range(len(ke)):
            for c in range(len(ke[0])):
                if ma_o(ten_ke, h, c) == ma:
                    return ten_ke, h, c
    return None


def phan_tich_ma_sp(noi_dung):
    """
    QR sản phẩm có dạng: K{kệ}-{Tên sản phẩm}-{Loại sản phẩm}-{STT}   (VD: K1-VFD-3P-01)
    Trả về (tên, loại, stt, số kệ) hoặc None nếu sai định dạng.
    (Vẫn đọc được tem cũ dạng Tên-Loại-STT — khi đó số kệ là None.)
    """
    ma = noi_dung.strip()
    m = re.fullmatch(r"[Kk](\d+)-(.+)-([^-]+)-(\d+)", ma)
    if m and m.group(2).strip() and m.group(3).strip() and int(m.group(4)) >= 1:
        return m.group(2).strip(), m.group(3).strip(), int(m.group(4)), int(m.group(1))
    phan = [p.strip() for p in ma.rsplit("-", 2)]
    if len(phan) != 3 or not all(phan) or not phan[2].isdigit() or int(phan[2]) < 1:
        return None
    return phan[0], phan[1], int(phan[2]), None


def hien_thong_bao_nhap(tb):
    """Hiển thị kết quả quét nhập hàng. Loại 'sai_ke' = khung ĐỎ lớn, nổi bật khi quét sản phẩm vào sai kệ."""
    loai, noi_dung = tb
    if loai == "sai_ke":
        html_nd = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", str(noi_dung).replace("<", "&lt;").replace(">", "&gt;"))
        html_nd = re.sub(r"`(.+?)`", r"<code style='color:#fff;background:#a31515;padding:1px 6px;border-radius:4px;'>\1</code>", html_nd)
        st.markdown(
            "<div style='background:#d32f2f;color:#fff;border:3px solid #7f0000;border-radius:10px;"
            "padding:16px 20px;margin:8px 0;font-size:1.05rem;line-height:1.5;'>"
            "<div style='font-size:1.4rem;font-weight:800;margin-bottom:4px;'>⛔ SẢN PHẨM NHẬP SAI KỆ!</div>"
            f"{html_nd}</div>", unsafe_allow_html=True)
        return
    {"success": st.success, "warning": st.warning, "error": st.error}[loai](noi_dung)


def xu_ly_quet_nhap(noi_dung):
    """
    Xử lý 1 lần quét (từ camera hoặc máy quét cầm tay):
      - Nếu là QR ô  -> chọn ô cần nhập (quét ô khác thì đổi sang ô đó)
      - Nếu là QR sản phẩm -> nhập vào vị trí trống đầu tiên của ô đã chọn
    Kết quả để lại trong st.session_state['nhap_tb'] = (loại, nội dung).
    """
    noi_dung = (noi_dung or "").strip()
    if not noi_dung:
        return

    # ---- Bước 1: QR vị trí ô ----
    o = phan_tich_ma_o(noi_dung)
    if o:
        st.session_state["nhap_o"] = o
        st.session_state["nhap_tb"] = (
            "success", f"Đã chọn ô **{ma_o(*o)}** ({o[0]}). Bây giờ hãy quét QR sản phẩm.")
        return

    if "nhap_o" not in st.session_state:
        sp_truoc = phan_tich_ma_sp(noi_dung)
        goi_y = f" Mã này thuộc **Kệ {sp_truoc[3]}**, hãy quét QR ô của Kệ {sp_truoc[3]}." if sp_truoc and sp_truoc[3] else ""
        st.session_state["nhap_tb"] = (
            "error", f"Chưa chọn ô. Hãy quét QR vị trí ô trước (đã quét: `{noi_dung}`).{goi_y}")
        return

    # ---- Bước 2: QR sản phẩm ----
    sp = phan_tich_ma_sp(noi_dung)
    if sp is None:
        st.session_state["nhap_tb"] = (
            "error", f"QR `{noi_dung}` không đúng định dạng **Kệ-Tên-Loại-STT** "
                     f"(VD: K1-VFD-3P-01) và cũng không phải mã ô.")
        return

    ten_hang, loai, stt, so_ke_qr = sp

    # Mã sản phẩm của kệ nào thì chỉ được nhập vào ô thuộc đúng kệ đó
    ten_ke_chon = st.session_state["nhap_o"][0]
    if so_ke_qr is None:
        st.session_state["nhap_tb"] = (
            "error", f"QR `{noi_dung}` chưa có số kệ nên không biết thuộc kệ nào. "
                     f"Hãy tạo lại tem theo dạng **K{{kệ}}-Tên-Loại-STT** (VD: K1-VFD-3P-01).")
        return
    if so_ke_qr != lay_so_ke(ten_ke_chon):
        st.session_state["nhap_tb"] = (
            "sai_ke", f"Sản phẩm `{noi_dung}` thuộc **Kệ {so_ke_qr}** nhưng ô đang chọn nằm ở **{ten_ke_chon}**. "
                     f"Hãy quét QR ô của Kệ {so_ke_qr}.")
        return

    # Chống nhập trùng trên TOÀN KHO: sản phẩm (tên + loại + STT) đang nằm ở bất kỳ ô nào thì không nhập nữa
    da_co = tim_san_pham_trong_kho(ten_hang, loai, stt)
    if da_co:
        cac_cho = ", ".join(f"ô **{ma_o(k, hh, cc)}** (Vị trí {i + 1})" for k, hh, cc, i in da_co)
        st.session_state["nhap_tb"] = (
            "warning", f"Sản phẩm `{noi_dung}` **đã có trong kho** tại {cac_cho}, không nhập lại. "
                       f"Nếu muốn chuyển sang ô khác, hãy xuất sản phẩm ra trước.")
        return

    ten_ke, h, c = st.session_state["nhap_o"]
    vi_tri_list = st.session_state.kho[ten_ke][h][c]
    vi_tri_trong = [i for i in range(len(vi_tri_list)) if vi_tri_list[i] is None]

    if not vi_tri_trong:
        st.session_state["nhap_tb"] = (
            "error", f"Ô **{ma_o(ten_ke, h, c)}** đã đầy {len(vi_tri_list)}/{len(vi_tri_list)} vị trí. "
                     f"Hãy quét QR của ô khác.")
        return

    idx = vi_tri_trong[0]
    vi_tri_list[idx] = {"ten_hang": ten_hang, "loai": loai, "stt": stt}
    ma_moi = ma_san_pham(ten_ke, h, c, ten_hang, loai, stt)
    luu_du_lieu(st.session_state.kho)
    ghi_lich_su("Nhập", ma_moi, ten_hang, loai)
    st.session_state["nhap_tb"] = (
        "success", f"Đã nhập `{ma_moi}` vào **Vị trí {idx + 1}** của ô {ma_o(ten_ke, h, c)}."
                   f"Có thể quét tiếp sản phẩm khác.")


# ============ CAMERA QUÉT QR TRỰC TIẾP (tự xây, không cần cài thêm thư viện) ============
# Khung camera hình vuông + khung quét hình vuông. Trình duyệt tự giải mã QR
# (dùng BarcodeDetector có sẵn của Chrome/Android, nếu không có thì dùng jsQR tải từ CDN).
QR_LIVE_HTML = r'''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  html, body { margin: 0; padding: 0; background: transparent; overflow: hidden;
               font-family: -apple-system, "Segoe UI", Roboto, sans-serif; }
  #box { position: relative; width: 100%; max-width: 420px; aspect-ratio: 1 / 1; margin: 0 auto;
         box-sizing: border-box; border: 3px solid #ffffff; border-radius: 16px;
         overflow: hidden; background: #000; transition: border-color .15s; }
  #box.ok { border-color: #21c55e; }
  video { position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; }
  #khung { position: absolute; top: 50%; left: 50%; width: 66%; aspect-ratio: 1 / 1;
           transform: translate(-50%, -50%); border-radius: 12px;
           box-shadow: 0 0 0 9999px rgba(0, 0, 0, .45); }
  .c { position: absolute; width: 28px; height: 28px; border: 4px solid #ffffff; transition: border-color .15s; }
  #box.ok .c { border-color: #21c55e; }
  .tl { top: 0; left: 0; border-right: 0; border-bottom: 0; border-top-left-radius: 12px; }
  .tr { top: 0; right: 0; border-left: 0; border-bottom: 0; border-top-right-radius: 12px; }
  .bl { bottom: 0; left: 0; border-right: 0; border-top: 0; border-bottom-left-radius: 12px; }
  .br { bottom: 0; right: 0; border-left: 0; border-top: 0; border-bottom-right-radius: 12px; }
  #line { position: absolute; left: 6%; right: 6%; height: 2px; background: #ffffff;
          box-shadow: 0 0 8px #ffffff; animation: quet 2.2s ease-in-out infinite alternate; }
  #box.ok #line { background: #21c55e; box-shadow: 0 0 8px #21c55e; }
  @keyframes quet { from { top: 8%; } to { top: 92%; } }
  #msg { position: absolute; left: 0; right: 0; bottom: 0; padding: 8px 10px; box-sizing: border-box;
         color: #fff; background: rgba(0, 0, 0, .65); font-size: 14px; text-align: center; }
  #msg.err { background: rgba(185, 28, 28, .9); }
</style>
</head>
<body>
<div id="box">
  <video id="video" playsinline muted autoplay></video>
  <div id="khung">
    <span class="c tl"></span><span class="c tr"></span><span class="c bl"></span><span class="c br"></span>
    <div id="line"></div>
  </div>
  <div id="msg">Đang mở camera...</div>
</div>
<script>
(function () {
  var COOLDOWN = 2500;           // ms: cùng 1 mã còn trong khung hình thì không gửi lại
  var JSQR_URL = "https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js";

  var box = document.getElementById("box");
  var video = document.getElementById("video");
  var msg = document.getElementById("msg");
  var canvas = document.createElement("canvas");
  var ctx = canvas.getContext("2d", { willReadFrequently: true });
  var detector = null, lastCode = null, lastSeen = 0, busy = false, stream = null, msgTimer = null;

  function send(type, extra) {
    window.parent.postMessage(Object.assign({ isStreamlitMessage: true, type: "streamlit:" + type }, extra), "*");
  }
  function setHeight() { send("setFrameHeight", { height: Math.ceil(box.getBoundingClientRect().height) + 2 }); }
  function setMsg(text, isErr) { msg.textContent = text; msg.className = isErr ? "err" : ""; }

  window.addEventListener("message", function (e) {
    if (e.data && e.data.type === "streamlit:render") setHeight();
  });
  if (window.ResizeObserver) new ResizeObserver(setHeight).observe(box);
  send("componentReady", { apiVersion: 1 });
  setHeight();

  function loadJsQR() {
    return new Promise(function (resolve, reject) {
      if (window.jsQR) return resolve();
      var s = document.createElement("script");
      s.src = JSQR_URL; s.onload = resolve; s.onerror = reject;
      document.head.appendChild(s);
    });
  }
  function needJsQR() {
    loadJsQR().catch(function () { setMsg("Không tải được bộ giải mã QR. Kiểm tra kết nối mạng.", true); });
  }

  if ("BarcodeDetector" in window) {
    try { detector = new window.BarcodeDetector({ formats: ["qr_code"] }); } catch (e) { detector = null; }
  }
  if (!detector) needJsQR();

  // Cắt đúng vùng vuông ở giữa khung hình (giống phần người dùng nhìn thấy) rồi giải mã
  async function decode() {
    var vw = video.videoWidth, vh = video.videoHeight;
    if (!vw || !vh) return null;
    var side = Math.min(vw, vh), out = Math.min(side, 720);
    canvas.width = out; canvas.height = out;
    ctx.drawImage(video, (vw - side) / 2, (vh - side) / 2, side, side, 0, 0, out, out);
    if (detector) {
      try {
        var r = await detector.detect(canvas);
        return (r && r.length) ? r[0].rawValue : null;
      } catch (e) { detector = null; needJsQR(); return null; }
    }
    if (window.jsQR) {
      var img = ctx.getImageData(0, 0, out, out);
      var res = window.jsQR(img.data, out, out, { inversionAttempts: "dontInvert" });
      if (res) {
        try {
          if (res.binaryData && res.binaryData.length)
            return new TextDecoder("utf-8").decode(new Uint8Array(res.binaryData));
        } catch (e) {}
        return res.data;
      }
    }
    return null;
  }

  function handle(code) {
    if (!code) return;
    code = code.trim();
    if (!code) return;
    var now = Date.now();
    if (code === lastCode && now - lastSeen < COOLDOWN) { lastSeen = now; return; }
    lastCode = code; lastSeen = now;
    box.classList.add("ok");
    setMsg("Đã quét: " + code, false);
    if (navigator.vibrate) { try { navigator.vibrate(80); } catch (e) {} }
    clearTimeout(msgTimer);
    msgTimer = setTimeout(function () { box.classList.remove("ok"); setMsg("Đưa mã QR vào khung vuông", false); }, 1500);
    send("setComponentValue", {
      value: { code: code, id: now + "-" + Math.random().toString(36).slice(2, 8) },
      dataType: "json"
    });
  }

  async function tick() {
    if (!busy && video.readyState >= 2) {
      busy = true;
      try { handle(await decode()); } catch (e) {}
      busy = false;
    }
    setTimeout(tick, 120);
  }

  async function startCam() {
    try {
      stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } },
        audio: false
      });
      video.srcObject = stream;
      await video.play();
      setMsg("Đưa mã QR vào khung vuông", false);
      tick();
    } catch (e) {
      setMsg("Không mở được camera (" + (e.name || e.message) + "). Hãy cho phép quyền camera và dùng HTTPS.", true);
    }
  }
  window.addEventListener("pagehide", function () {
    if (stream) stream.getTracks().forEach(function (t) { t.stop(); });
  });
  startCam();
})();
</script>
</body>
</html>
'''


def _tao_component_qr():
    """Ghi file giao diện camera ra thư mục rồi đăng ký làm component của Streamlit."""
    for thu_muc in (Path(__file__).resolve().parent / "qr_live_component",
                    Path(tempfile.gettempdir()) / "qr_live_component"):
        try:
            thu_muc.mkdir(exist_ok=True)
            f = thu_muc / "index.html"
            if not f.exists() or f.read_text(encoding="utf-8") != QR_LIVE_HTML:
                f.write_text(QR_LIVE_HTML, encoding="utf-8")
            return components.declare_component("qr_live_scanner", path=str(thu_muc))
        except OSError:
            continue
    return None


_QR_LIVE = _tao_component_qr()


# ============ CHỨC NĂNG: NHẬP HÀNG (QUÉT QR) ============
# Gồm 2 trang:
#   - "nhap"      : trang tổng quan (xem ô đang chọn, các vị trí, kết quả quét gần nhất)
#   - "quet_nhap" : trang camera riêng, khung quét hình vuông

def nhap_bang_quet_qr():
    """Tab quét QR của trang nhập hàng (tổng quan + nút mở camera)."""
    if st.button("Mở camera quét QR", type="primary", use_container_width=True):
        st.session_state.pop("nhap_qr_cuoi", None)
        st.session_state["nhap_cam_n"] = st.session_state.get("nhap_cam_n", 0) + 1  # camera mới
        st.session_state.chuc_nang = "quet_nhap"
        st.rerun()

    khung_tinh_trang_kho("nhap")
    tb = st.session_state.get("nhap_tb")
    if tb and tb[0] == "sai_ke":                      # trang này chỉ hiện khung đỏ khi quét sai kệ
        hien_thong_bao_nhap(tb)

    o_chon = st.session_state.get("nhap_o")
    if o_chon is not None:
        ten_ke, h, c = o_chon
        so_trong = sum(1 for vt in st.session_state.kho[ten_ke][h][c] if vt is None)
        hien_thi_danh_sach_ma_sp(ten_ke, h, c)
        if st.button("Bỏ chọn ô này"):
            st.session_state.pop("nhap_o", None)
            st.rerun()


def man_hinh_nhap():
    """Trang nhập hàng: 2 tab — quét QR hoặc nhập tay."""
    tieu_de_trang("Nhập hàng vào kho")
    tab_quet, tab_tay = st.tabs(["Quét QR", "Nhập tay"])
    with tab_quet:
        nhap_bang_quet_qr()
    with tab_tay:
        nhap_tay()
    nut_quay_lai()


def nhap_tay():
    """Nhập hàng bằng tay: chọn ô trên sơ đồ kệ -> nhập tên, loại, số lượng (STT tự đánh)."""
    tb = st.session_state.pop("nhap_tay_tb", None)    # kết quả của lần nhập trước
    if tb:
        st.success(tb)

    nhan_muc("Chọn ô cần nhập")
    ten_ke = st.selectbox("Chọn kệ", list(st.session_state.kho), key="nhap_tay_ke")
    o_chon = ve_so_do_ke(ten_ke, key_prefix="nhap")
    if o_chon:
        st.session_state["nhap_ke"] = ten_ke
        st.session_state["nhap_hang"] = o_chon[0]
        st.session_state["nhap_cot"] = o_chon[1]

    # Chỉ hiển thị ô đã chọn nếu nó thuộc kệ đang xem
    if "nhap_hang" in st.session_state and st.session_state.get("nhap_ke") == ten_ke:
        h, c = st.session_state["nhap_hang"], st.session_state["nhap_cot"]
        st.divider()
        hien_thi_danh_sach_ma_sp(ten_ke, h, c)

        vi_tri_list = st.session_state.kho[ten_ke][h][c]
        vi_tri_trong = [i for i in range(len(vi_tri_list)) if vi_tri_list[i] is None]
        ten_hien_co = sorted({vt["ten_hang"] for vt in vi_tri_list if vt is not None})

        st.divider()
        st.subheader("Nhập sản phẩm")

        if not vi_tri_trong:
            st.error(f"Ô này đã đầy {len(vi_tri_list)}/{len(vi_tri_list)} vị trí, không thể nhập thêm. "
                     f"Vui lòng chọn ô khác hoặc xuất bớt.")
        else:
            tuy_chon_ten = ten_hien_co + ["+ Thêm tên sản phẩm mới"]
            lua_chon_ten = st.selectbox("Chọn tên sản phẩm", tuy_chon_ten, key="chon_ten_nhap")

            with st.form("form_nhap"):
                if lua_chon_ten == "+ Thêm tên sản phẩm mới":
                    ten_hang = st.text_input("Tên sản phẩm mới")
                else:
                    ten_hang = lua_chon_ten
                    st.text_input("Tên sản phẩm", value=ten_hang, disabled=True)

                loai = st.text_input("Loại sản phẩm", placeholder="VD: Cotton, Da bò, Size M ...")
                so_luong_tao = st.number_input(
                    f"Số sản phẩm cần nhập (mỗi sản phẩm chiếm 1 vị trí, tối đa {len(vi_tri_trong)})",
                    min_value=1, max_value=len(vi_tri_trong), value=1, step=1
                )
                submit = st.form_submit_button("Xác nhận nhập hàng", type="primary")

                if submit:
                    if not ten_hang or not ten_hang.strip():
                        st.error("Vui lòng nhập tên sản phẩm.")
                    elif not loai or not loai.strip():
                        st.error("Vui lòng nhập loại sản phẩm.")
                    else:
                        ten_hang = ten_hang.strip()
                        loai = loai.strip()
                        # STT lớn nhất của đúng cặp (tên, loại) trên TOÀN KHO (so sánh theo dạng chuẩn hóa)
                        # để sản phẩm nhập ở ô khác không bị trùng STT với sản phẩm đã có
                        stt_hien_tai = stt_lon_nhat_trong_kho(ten_hang, loai)

                        ma_moi = []
                        for i in range(int(so_luong_tao)):
                            idx_vi_tri = vi_tri_trong[i]
                            stt = stt_hien_tai + i + 1
                            vi_tri_list[idx_vi_tri] = {"ten_hang": ten_hang, "loai": loai, "stt": stt}
                            ma_moi.append(ma_san_pham(ten_ke, h, c, ten_hang, loai, stt))

                        luu_du_lieu(st.session_state.kho)
                        for ma_sp in ma_moi:
                            ghi_lich_su("Nhập", ma_sp, ten_hang, loai)

                        st.session_state["nhap_tay_tb"] = (
                            f"Đã nhập {len(ma_moi)} '{ten_hang}' (loại: {loai}) vào ô {ma_o(ten_ke, h, c)}. Mã: "
                            + ", ".join(f"`{m}`" for m in ma_moi))
                        for k in ("nhap_hang", "nhap_cot", "nhap_ke"):
                            st.session_state.pop(k, None)
                        st.rerun()


def man_hinh_quet_nhap():
    """Trang camera riêng: quét QR trực tiếp, khung hình vuông."""
    if st.button("Quay lại trang nhập hàng"):
        st.session_state.chuc_nang = "nhap"
        st.rerun()
    tieu_de_trang("Quét QR nhập hàng", nho=True)

    if _QR_LIVE is None:
        st.error("Không tạo được thành phần camera (không ghi được file vào thư mục chạy app). "
                 "Hãy kiểm tra quyền ghi thư mục.")
        return

    # --- Camera quét liên tục (đặt trước các phần động để camera không bị tải lại) ---
    n = st.session_state.get("nhap_cam_n", 0)
    ket_qua = _QR_LIVE(key=f"nhap_qr_{n}", default=None)

    # Mỗi lần quét có 1 mã 'id' riêng; chỉ xử lý khi là lần quét MỚI
    if ket_qua and ket_qua.get("id") != st.session_state.get("nhap_qr_cuoi"):
        st.session_state["nhap_qr_cuoi"] = ket_qua["id"]
        xu_ly_quet_nhap(str(ket_qua.get("code", "")))

    # --- Trạng thái hiện tại ---
    o_chon = st.session_state.get("nhap_o")
    if o_chon is None:
        st.subheader("Bước 1: Quét QR vị trí ô")
    else:
        ten_ke, h, c = o_chon
        so_trong = sum(1 for vt in st.session_state.kho[ten_ke][h][c] if vt is None)
        st.subheader("Bước 2: Quét QR sản phẩm")

    tb = st.session_state.get("nhap_tb")
    if tb:
        hien_thong_bao_nhap(tb)

    col_a, col_b = st.columns(2)
    with col_a:
        if o_chon is not None and st.button("Chọn ô khác", use_container_width=True):
            st.session_state.pop("nhap_o", None)
            st.session_state.pop("nhap_qr_cuoi", None)
            st.session_state["nhap_cam_n"] = n + 1   # làm mới camera để quét lại được mã cũ
            st.rerun()
    with col_b:
        if st.button("Xong, xem kết quả", use_container_width=True):
            st.session_state.chuc_nang = "nhap"
            st.rerun()


# ============ CHỨC NĂNG: XUẤT HÀNG (QUÉT QR SẢN PHẨM) ============
# Gồm 2 trang:
#   - "xuat"      : trang tổng quan (kết quả quét gần nhất, danh sách đã xuất trong phiên)
#   - "quet_xuat" : trang camera riêng, chỉ cần quét QR sản phẩm (Kệ-Tên-Loại-STT) là xuất

def tim_san_pham_trong_kho(ten_hang, loai, stt):
    """Tìm mọi vị trí trong kho đang chứa sản phẩm này. Trả về list (tên kệ, tầng, ô, vị trí)."""
    ket_qua = []
    for ten_ke, ke in st.session_state.kho.items():
        for h in range(len(ke)):
            for c in range(len(ke[0])):
                for i, vt in enumerate(ke[h][c]):
                    if (vt is not None and chuan_hoa(vt["ten_hang"]) == chuan_hoa(ten_hang)
                            and chuan_hoa(vt["loai"]) == chuan_hoa(loai) and vt["stt"] == stt):
                        ket_qua.append((ten_ke, h, c, i))
    return ket_qua


def stt_lon_nhat_trong_kho(ten_hang, loai):
    """STT lớn nhất của cặp (tên, loại) đang có trong toàn kho (0 nếu chưa có)."""
    stt_max = 0
    for ten_ke, ke in st.session_state.kho.items():
        for h in range(len(ke)):
            for c in range(len(ke[0])):
                for vt in ke[h][c]:
                    if (vt is not None and chuan_hoa(vt["ten_hang"]) == chuan_hoa(ten_hang)
                            and chuan_hoa(vt["loai"]) == chuan_hoa(loai)):
                        stt_max = max(stt_max, vt["stt"])
    return stt_max


def thuc_hien_xuat(ten_ke, h, c, idx):
    """Xuất sản phẩm ở 1 vị trí: trả vị trí về trống, lưu kho, ghi lịch sử, ghi vào danh sách phiên."""
    vi_tri_list = st.session_state.kho[ten_ke][h][c]
    vt = vi_tri_list[idx]
    ma = ma_san_pham(ten_ke, h, c, vt["ten_hang"], vt["loai"], vt["stt"])
    vi_tri_list[idx] = None
    luu_du_lieu(st.session_state.kho)
    ghi_lich_su("Xuất", ma, vt["ten_hang"], vt["loai"])
    st.session_state.setdefault("xuat_ds", []).append({
        "thoi_gian": datetime.now().strftime("%H:%M:%S"),
        "ma": ma, "ten_hang": vt["ten_hang"], "loai": vt["loai"],
        "vi_tri": f"Ô {ma_o(ten_ke, h, c)} - Vị trí {idx + 1}",
    })
    st.session_state["xuat_tb"] = (
        "success", f"Đã xuất `{ma}` ({ten_day_du(vt['ten_hang'])} / {ten_loai_day_du(vt['ten_hang'], vt['loai'])}) "
                   f"từ ô **{ma_o(ten_ke, h, c)}** — Vị trí {idx + 1}.")


def xu_ly_quet_xuat(noi_dung):
    """Xử lý 1 lần quét ở chức năng xuất: tìm sản phẩm trong kho theo QR rồi xuất."""
    noi_dung = (noi_dung or "").strip()
    if not noi_dung:
        return
    st.session_state.pop("xuat_cho_chon", None)   # lần quét mới thì bỏ lựa chọn đang chờ

    if phan_tich_ma_o(noi_dung):
        st.session_state["xuat_tb"] = (
            "error", f"`{noi_dung}` là mã của ô. Xuất hàng chỉ cần quét QR **sản phẩm** (dạng Kệ-Tên-Loại-STT).")
        return

    sp = phan_tich_ma_sp(noi_dung)
    if sp is None:
        st.session_state["xuat_tb"] = (
            "error", f"QR `{noi_dung}` không đúng định dạng **Kệ-Tên-Loại-STT** (VD: K1-VFD-3P-01).")
        return

    ten_hang, loai, stt, _so_ke = sp
    vi_tri_tim_thay = tim_san_pham_trong_kho(ten_hang, loai, stt)

    if not vi_tri_tim_thay:
        st.session_state["xuat_tb"] = (
            "warning", f"Không tìm thấy sản phẩm `{noi_dung}` trong kho (có thể đã được xuất rồi).")
    elif len(vi_tri_tim_thay) == 1:
        thuc_hien_xuat(*vi_tri_tim_thay[0])
    else:
        # Cùng 1 mã nằm ở nhiều ô -> không tự đoán, để người dùng chọn xuất từ ô nào
        st.session_state["xuat_cho_chon"] = {"ma": noi_dung, "vi_tri": vi_tri_tim_thay}
        st.session_state["xuat_tb"] = (
            "warning", f"Sản phẩm `{noi_dung}` có ở **{len(vi_tri_tim_thay)} vị trí** khác nhau. "
                       f"Hãy chọn xuất từ vị trí nào bên dưới.")


def _hien_thi_ds_da_xuat():
    ds = st.session_state.get("xuat_ds", [])
    if not ds:
        return
    st.markdown(f"##### Đã xuất trong phiên này: **{len(ds)}** sản phẩm")
    bang_ten, bang_loai = doc_ten_day_du(), doc_ten_loai_day_du()
    for m in reversed(ds):
        st.write(f"`{m['thoi_gian']}` — `{m['ma']}` — {ten_day_du(m['ten_hang'], bang_ten)} / "
                 f"{ten_loai_day_du(m['ten_hang'], m['loai'], bang_loai)} — {m['vi_tri']}")


def xuat_bang_quet_qr():
    """Tab quét QR của trang xuất hàng (tổng quan + nút mở camera)."""

    if st.button("Mở camera quét QR", type="primary", use_container_width=True):
        st.session_state.pop("xuat_qr_cuoi", None)
        st.session_state["xuat_cam_n"] = st.session_state.get("xuat_cam_n", 0) + 1  # camera mới
        st.session_state.chuc_nang = "quet_xuat"
        st.rerun()

    khung_tinh_trang_kho("xuat")
    tb = st.session_state.get("xuat_tb")
    if tb:
        nhan_muc("Kết quả lần quét gần nhất")
        {"success": st.success, "warning": st.warning, "error": st.error}[tb[0]](tb[1])
    _hien_thi_ds_da_xuat()


def man_hinh_xuat():
    """Trang xuất hàng: 2 tab — quét QR hoặc xuất tay."""
    tieu_de_trang("Xuất hàng khỏi kho")
    tab_quet, tab_tay = st.tabs(["Quét QR", "Xuất tay"])
    with tab_quet:
        xuat_bang_quet_qr()
    with tab_tay:
        xuat_tay()
    nut_quay_lai()


def xuat_tay():
    """Xuất hàng bằng tay: chọn ô đang có hàng -> chọn (các) mã sản phẩm cần xuất."""

    tb = st.session_state.pop("xuat_tay_tb", None)    # kết quả của lần xuất trước
    if tb:
        st.success(tb)

    nhan_muc("Chọn ô cần xuất")
    ten_ke = st.selectbox("Chọn kệ", list(st.session_state.kho), key="xuat_tay_ke")
    o_chon = ve_so_do_ke(ten_ke, key_prefix="xuat")
    if o_chon:
        h, c = o_chon
        if all(vt is None for vt in st.session_state.kho[ten_ke][h][c]):
            st.warning("Ô bạn vừa chọn đang trống, không có hàng để xuất.")
        else:
            st.session_state["xuat_ke"] = ten_ke
            st.session_state["xuat_hang"] = h
            st.session_state["xuat_cot"] = c

    # Chỉ hiển thị ô đã chọn nếu nó thuộc kệ đang xem
    if "xuat_hang" in st.session_state and st.session_state.get("xuat_ke") == ten_ke:
        h, c = st.session_state["xuat_hang"], st.session_state["xuat_cot"]
        vi_tri_list = st.session_state.kho[ten_ke][h][c]

        st.divider()
        hien_thi_danh_sach_ma_sp(ten_ke, h, c)

        vi_tri_co_hang = [i for i in range(len(vi_tri_list)) if vi_tri_list[i] is not None]

        if not vi_tri_co_hang:
            pass
        else:
            st.divider()
            bang_ten, bang_loai = doc_ten_day_du(), doc_ten_loai_day_du()
            tuy_chon = [
                f"Vị trí {i + 1}: "
                f"{ma_san_pham(ten_ke, h, c, vi_tri_list[i]['ten_hang'], vi_tri_list[i]['loai'], vi_tri_list[i]['stt'])} "
                f"— {ten_day_du(vi_tri_list[i]['ten_hang'], bang_ten)} / {ten_loai_day_du(vi_tri_list[i]['ten_hang'], vi_tri_list[i]['loai'], bang_loai)}"
                for i in vi_tri_co_hang
            ]
            lua_chon = st.multiselect("Chọn (một hoặc nhiều) mã sản phẩm cần xuất", tuy_chon,
                                      key="xuat_tay_chon")

            if lua_chon and st.button("Xác nhận xuất hàng", type="primary", key="xuat_tay_ok"):
                idx_can_xuat = [vi_tri_co_hang[tuy_chon.index(lc)] for lc in lua_chon]
                da_xuat = []
                for idx in idx_can_xuat:
                    vt = vi_tri_list[idx]
                    ma = ma_san_pham(ten_ke, h, c, vt["ten_hang"], vt["loai"], vt["stt"])
                    ghi_lich_su("Xuất", ma, vt["ten_hang"], vt["loai"])
                    da_xuat.append(ma)
                    vi_tri_list[idx] = None  # trả lại vị trí trống

                luu_du_lieu(st.session_state.kho)
                st.session_state["xuat_tay_tb"] = (
                    f"Đã xuất {len(da_xuat)} sản phẩm: " + ", ".join(f"`{m}`" for m in da_xuat))
                st.session_state.pop("xuat_tay_chon", None)   # chỉ bỏ các mã đã tick; GIỮ nguyên ô đang chọn để xuất tiếp
                st.rerun()


def man_hinh_quet_xuat():
    """Trang camera riêng cho xuất hàng (khung vuông, tự nhận diện)."""
    if st.button("Quay lại trang xuất hàng"):
        st.session_state.chuc_nang = "xuat"
        st.rerun()
    tieu_de_trang("Quét QR xuất hàng", nho=True)

    if _QR_LIVE is None:
        st.error("Không tạo được thành phần camera (không ghi được file vào thư mục chạy app). "
                 "Hãy kiểm tra quyền ghi thư mục.")
        return

    # --- Camera quét liên tục (đặt trước các phần động để camera không bị tải lại) ---
    n = st.session_state.get("xuat_cam_n", 0)
    ket_qua = _QR_LIVE(key=f"xuat_qr_{n}", default=None)

    if ket_qua and ket_qua.get("id") != st.session_state.get("xuat_qr_cuoi"):
        st.session_state["xuat_qr_cuoi"] = ket_qua["id"]
        xu_ly_quet_xuat(str(ket_qua.get("code", "")))

    st.subheader("Quét QR sản phẩm cần xuất")
    tb = st.session_state.get("xuat_tb")
    if tb:
        {"success": st.success, "warning": st.warning, "error": st.error}[tb[0]](tb[1])

    # Trường hợp cùng 1 mã có ở nhiều vị trí: chọn vị trí cần xuất
    cho = st.session_state.get("xuat_cho_chon")
    if cho:
        for k, (ten_ke, h, c, i) in enumerate(cho["vi_tri"]):
            if st.button(f"Xuất từ ô {ma_o(ten_ke, h, c)} — Vị trí {i + 1}",
                         key=f"xuat_chon_{k}", use_container_width=True):
                thuc_hien_xuat(ten_ke, h, c, i)
                st.session_state.pop("xuat_cho_chon", None)
                st.rerun()

    st.divider()
    _hien_thi_ds_da_xuat()
    if st.button("Xong, xem kết quả", use_container_width=True):
        st.session_state.chuc_nang = "xuat"
        st.rerun()


# ============ CHỨC NĂNG: LỊCH SỬ NHẬP / XUẤT ============
def man_hinh_lich_su():
    tieu_de_trang("Lịch sử nhập / xuất")

    du_lieu = doc_lich_su()
    if not du_lieu:
        nut_quay_lai()
        return

    df = pd.DataFrame(du_lieu).rename(columns={
        "thoi_gian": "Thời gian", "hanh_dong": "Hành động", "ma_san_pham": "Mã sản phẩm",
        "ten_hang": "Tên hàng", "loai": "Loại"
    })
    bang_ten, bang_loai = doc_ten_day_du(), doc_ten_loai_day_du()
    df["Loại"] = [ten_loai_day_du(t, l, bang_loai) if isinstance(t, str) and isinstance(l, str) else l
                  for t, l in zip(df["Tên hàng"], df["Loại"])]       # phải đổi Loại TRƯỚC khi đổi Tên hàng (cần mã sản phẩm gốc)
    df["Tên hàng"] = df["Tên hàng"].map(lambda t: ten_day_du(t, bang_ten) if isinstance(t, str) else t)
    df["Thời gian"] = pd.to_datetime(df["Thời gian"], errors="coerce")
    df["Mã ô"] = df["Mã sản phẩm"].str.split("-").str[0]      # K1T101-AOTHUN-... -> K1T101
    df = df.sort_values("Thời gian", ascending=False).reset_index(drop=True)

    # ---- Bộ lọc ----
    nhan_muc("Bộ lọc")
    c1, c2, c3, c4 = st.columns([1, 1.4, 1, 1.4])
    hanh_dong = c1.selectbox("Hành động", ["Tất cả", "Nhập", "Xuất"])
    ngay_min, ngay_max = df["Thời gian"].min().date(), df["Thời gian"].max().date()
    khoang_ngay = c2.date_input("Khoảng ngày", value=(ngay_min, ngay_max),
                                min_value=ngay_min, max_value=ngay_max)
    o_loc = c3.selectbox("Ô", ["Tất cả"] + sorted(df["Mã ô"].dropna().unique()))
    tu_khoa = c4.text_input("Tìm (mã / tên / loại)", placeholder="VD: AOTHUN, cotton ...")

    loc = df
    if hanh_dong != "Tất cả":
        loc = loc[loc["Hành động"] == hanh_dong]
    if isinstance(khoang_ngay, tuple) and len(khoang_ngay) == 2:
        tu, den = khoang_ngay
        loc = loc[(loc["Thời gian"].dt.date >= tu) & (loc["Thời gian"].dt.date <= den)]
    if o_loc != "Tất cả":
        loc = loc[loc["Mã ô"] == o_loc]
    if tu_khoa.strip():
        k = tu_khoa.strip().lower()
        loc = loc[loc[["Mã sản phẩm", "Tên hàng", "Loại"]].astype(str)
                  .apply(lambda col: col.str.lower().str.contains(k, regex=False)).any(axis=1)]

    # ---- Thống kê + bảng ----
    nhan_muc("Kết quả")
    st.markdown(html_the_so_lieu([
        ("", len(loc), "Số dòng hiển thị", "theo bộ lọc hiện tại", "blue"),
        ("", int((loc["Hành động"] == "Nhập").sum()), "Lượt nhập", "trong kết quả lọc", "green"),
        ("", int((loc["Hành động"] == "Xuất").sum()), "Lượt xuất", "trong kết quả lọc", "orange"),
    ]), unsafe_allow_html=True)

    hien_thi = loc[["Thời gian", "Hành động", "Mã sản phẩm", "Mã ô", "Tên hàng", "Loại"]].copy()
    hien_thi["Thời gian"] = hien_thi["Thời gian"].dt.strftime("%Y-%m-%d %H:%M:%S")
    st.dataframe(hien_thi, use_container_width=True, hide_index=True)

    st.download_button("Tải CSV (đang lọc)", hien_thi.to_csv(index=False).encode("utf-8-sig"),
                       file_name="lich_su_kho.csv", mime="text/csv")

    nut_quay_lai()


# ============ CHỨC NĂNG: TẠO MÃ QR (TAB 1: SẢN PHẨM, TAB 2: VỊ TRÍ Ô) -> XUẤT FILE WORD ============
# Mỗi mã QR có nội dung  Kệ-Tên-Loại-STT  (VD: K1-VFD-3P-01) — đúng định dạng mà các chức năng
# Nhập / Xuất / Kiểm tra đang quét. Thư viện: pip install python-docx qrcode
# (nếu chưa cài qrcode nhưng có OpenCV thì app tự dùng OpenCV để vẽ mã QR).
MAX_MA_QR = 1000          # tối đa số mã QR tạo trong 1 lần


def tao_anh_qr_png(noi_dung):
    """Vẽ 1 mã QR (nội dung UTF-8, tiếng Việt có dấu OK) và trả về bytes ảnh PNG."""
    try:
        import qrcode
    except ImportError:
        qrcode = None

    if qrcode is not None:
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4)      # viền trắng 4 module theo chuẩn QR
        qr.add_data(noi_dung)
        qr.make(fit=True)
        anh = qr.make_image(fill_color="black", back_color="white")
        bo_dem = io.BytesIO()
        anh.save(bo_dem)
        return bo_dem.getvalue()

    try:                                   # phương án dự phòng: OpenCV
        import cv2
    except ImportError:
        raise RuntimeError("Chưa cài thư viện tạo mã QR. Hãy chạy: pip install qrcode")
    ma_tran = cv2.QRCodeEncoder.create().encode(noi_dung)
    ma_tran = cv2.copyMakeBorder(ma_tran, 4, 4, 4, 4, cv2.BORDER_CONSTANT, value=255)   # viền trắng 4 module theo chuẩn QR
    ti_le = max(1, 400 // ma_tran.shape[0])
    ma_tran = cv2.resize(ma_tran, None, fx=ti_le, fy=ti_le, interpolation=cv2.INTER_NEAREST)
    return cv2.imencode(".png", ma_tran)[1].tobytes()


def _chuoi(v):
    return "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v).strip()


def _so_nguyen(v):
    try:
        return None if v is None or pd.isna(v) else int(v)
    except (TypeError, ValueError):
        return None


def kiem_tra_bang_qr(cac_dong):
    """
    Kiểm tra bảng nhập (mỗi dòng: kệ, tên sản phẩm, mã sản phẩm, tên loại, mã loại, STT bắt đầu, số lượng) và sinh danh sách mã dạng K1-VFD-3P-01.
    Trả về (danh_sach_ma, danh_sach_loi). Dòng hoàn toàn trống được bỏ qua.
    """
    ds, loi = [], []
    ten_loai_cua = {}                                            # (mã sản phẩm, mã loại) chuẩn hóa -> tên loại
    ten_dd_cua = {}                                              # mã sản phẩm (chuẩn hóa) -> tên sản phẩm, để bắt xung đột trong bảng
    for i, d in enumerate(cac_dong, start=1):
        ten, loai = _chuoi(d.get("Mã sản phẩm")), _chuoi(d.get("Mã loại sản phẩm"))
        ten_loai_dd = _chuoi(d.get("Tên loại sản phẩm"))
        ten_dd = _chuoi(d.get("Tên sản phẩm"))
        so_ke = _so_nguyen(d.get("Kệ"))
        bat_dau, so_luong = _so_nguyen(d.get("STT bắt đầu")), _so_nguyen(d.get("Số lượng"))
        if not ten and not ten_dd and not loai and not ten_loai_dd and so_luong is None:
            continue
        if so_ke is None or so_ke < 1:
            loi.append(f"Dòng {i}: thiếu số kệ (VD: 1 cho Kệ 1)."); continue
        if st.session_state.get("kho") and not any(lay_so_ke(k) == so_ke for k in st.session_state.kho):
            loi.append(f"Dòng {i}: không có Kệ {so_ke} trong kho."); continue
        if not ten:
            loi.append(f"Dòng {i}: thiếu mã sản phẩm (phần này nằm trong mã QR, VD: VFD)."); continue
        if not loai:
            loi.append(f"Dòng {i}: thiếu mã loại sản phẩm (phần này nằm trong mã QR, VD: 3P)."); continue
        if not chuan_hoa(ten) or not chuan_hoa(loai):
            loi.append(f"Dòng {i}: mã sản phẩm và mã loại phải có ít nhất 1 chữ hoặc số."); continue
        ten_dd = ten_dd or ten                                   # bỏ trống tên sản phẩm -> hiển thị luôn mã sản phẩm
        cu = ten_dd_cua.get(chuan_hoa(ten))
        if cu is not None and cu != ten_dd:
            loi.append(f"Dòng {i}: mã sản phẩm '{ten}' đã được dùng ở dòng trên với tên sản phẩm khác ('{cu}'). "
                       f"Mỗi mã sản phẩm chỉ có 1 tên sản phẩm."); continue
        ten_dd_cua[chuan_hoa(ten)] = ten_dd
        ten_loai_dd = ten_loai_dd or loai                        # bỏ trống tên loại -> hiển thị luôn mã loại
        khoa_loai = (chuan_hoa(ten), chuan_hoa(loai))
        cu_loai = ten_loai_cua.get(khoa_loai)
        if cu_loai is not None and cu_loai != ten_loai_dd:
            loi.append(f"Dòng {i}: mã loại '{loai}' của mã sản phẩm '{ten}' đã được dùng ở dòng trên với tên loại khác "
                       f"('{cu_loai}'). Mỗi mã loại chỉ có 1 tên loại."); continue
        ten_loai_cua[khoa_loai] = ten_loai_dd
        bat_dau = 1 if bat_dau is None else bat_dau
        if bat_dau < 1:
            loi.append(f"Dòng {i}: STT bắt đầu phải từ 1 trở lên."); continue
        if so_luong is None or so_luong < 1:
            loi.append(f"Dòng {i}: số lượng phải từ 1 trở lên."); continue
        for stt in range(bat_dau, bat_dau + so_luong):
            ds.append({"ke": so_ke, "ten": ten, "ten_day_du": ten_dd, "loai": loai, "ten_loai_day_du": ten_loai_dd, "stt": stt,
                       "noi_dung": ma_san_pham_tu_so_ke(so_ke, ten, loai, stt)})
    if not loi and not ds:
        loi.append("Chưa có sản phẩm nào. Hãy nhập tên, loại, STT bắt đầu và số lượng.")
    if not loi and len(ds) > MAX_MA_QR:
        loi.append(f"Tổng cộng {len(ds)} mã, vượt giới hạn {MAX_MA_QR} mã mỗi lần. Hãy chia thành nhiều lần tạo.")
    return ds, loi


KHOANG_KHUNG_CM = 0.5      # khung bao quanh mã QR cách chấm đen của mã QR 5 mm


def so_module_qr(noi_dung):
    """Số module (ô vuông nhỏ) mỗi cạnh của mã QR, cùng cách mã hóa với tao_anh_qr_png."""
    try:
        import qrcode
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4)
        qr.add_data(noi_dung)
        qr.make(fit=True)
        return qr.modules_count
    except ImportError:
        pass
    import cv2
    return cv2.QRCodeEncoder.create().encode(noi_dung).shape[0]


def tao_anh_qr_cat_png(noi_dung):
    """Ảnh PNG mã QR đã cắt sát chấm đen (bỏ viền trắng sẵn có); khoảng trắng quanh QR do khung tem đảm nhiệm."""
    from PIL import Image
    goc = Image.open(io.BytesIO(tao_anh_qr_png(noi_dung))).convert("RGB")
    n = so_module_qr(noi_dung)
    le = round(4 * goc.width / (n + 8))                    # viền trắng 4 module sẵn có của ảnh QR
    bo_dem = io.BytesIO()
    goc.crop((le, le, goc.width - le, goc.height - le)).save(bo_dem, format="PNG")
    return bo_dem.getvalue()


def tinh_bo_cuc_tem(canh_khung_cm):
    """Từ cạnh khung vuông của tem (cm) -> (số tem mỗi hàng, bề rộng mỗi ô tem cm) trên khổ A4 dọc, lề 1.5 cm."""
    rong_dung_duoc = 21.0 - 2 * 1.5
    rong_toi_thieu = max(canh_khung_cm + 0.5, 2.8)         # + 0.5 cm: khe giữa khung đen và đường cắt nét đứt
    so_cot = max(1, int(rong_dung_duoc // rong_toi_thieu))
    return so_cot, rong_dung_duoc / so_cot


def tao_file_word_qr(ds_ma, kich_thuoc_cm=2.0, tieu_de="Tem QR sản phẩm"):
    """Tạo file Word (A4 dọc): mỗi tem gồm mã QR + nội dung mã bên dưới, cả hai nằm trong 1 khung liền nét cách nội dung 5 mm.
    Mã QR in ra to đúng cỡ `kich_thuoc_cm` (kể cả viền trắng chuẩn). Số tem mỗi hàng tự tính theo khổ giấy. Trả về bytes .docx."""
    from docx import Document
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor, Twips

    LE, RONG_TRANG = 1.5, 21.0
    kich_thuoc_qr = kich_thuoc_cm                         # cạnh ảnh QR in ra (hình vuông, gồm viền trắng chuẩn)
    co_chu = 8 if kich_thuoc_cm < 3 else 9                # cỡ chữ dưới QR (pt)
    n_max = max(so_module_qr(m["noi_dung"]) for m in ds_ma)
    rong_qr_max = kich_thuoc_qr * n_max / (n_max + 8)     # bề rộng chấm đen của mã nhiều module nhất
    rong_chu = max(len(m["noi_dung"]) for m in ds_ma) * co_chu * 0.6 / 72 * 2.54       # ước lượng bề rộng dòng chữ dài nhất (cm)
    # Khung HÌNH VUÔNG cố định cho cả file: cạnh = max(bề rộng cần, chiều cao cần) của nội dung + 2 lần khoảng cách 5 mm
    rong_can = max(rong_qr_max, rong_chu) + 2 * KHOANG_KHUNG_CM
    cao_can = rong_qr_max + (3 + co_chu * 1.2) / 72 * 2.54 + 2 * KHOANG_KHUNG_CM    # QR + khe 3pt + 1 dòng chữ
    cao_noi_dung = rong_qr_max + (3 + co_chu * 1.2) / 72 * 2.54       # QR + khe 3pt + 1 dòng chữ (cm)
    can_toi_thieu = max(rong_chu, cao_noi_dung) + 0.3                  # chừa tối thiểu 1,5 mm mỗi phía
    rong_khung = max(2.5, math.ceil(can_toi_thieu * 2) / 2)            # cạnh khung vuông (cm): mặc định 2,5 cm; nếu nội dung to hơn thì làm tròn lên bội số 0,5 cm
    so_cot, rong_cot = tinh_bo_cuc_tem(rong_khung)        # số tem mỗi hàng & bề rộng mỗi ô tem (cm)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(RONG_TRANG), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(LE)
    sec.top_margin = sec.bottom_margin = Cm(LE)
    sec.header_distance = Cm(0.7)

    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Arial", Pt(9)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    dau_trang = sec.header.paragraphs[0]
    chay = dau_trang.add_run(f"{tieu_de} — tạo ngày {datetime.now().strftime('%d/%m/%Y')} — {len(ds_ma)} mã")
    chay.font.size, chay.font.color.rgb = Pt(8), RGBColor(0x80, 0x80, 0x80)

    # Mỗi tem = khối 3x3 ô trong MỘT bảng duy nhất (không lồng bảng, vì Word không giữ chiều cao cố định của bảng lồng):
    #   ô giữa = khung vuông viền đen liền nét (cạnh = rong_khung), 8 ô quanh là khe trắng; rìa ngoài khối là nét đứt để cắt tem.
    so_hang = math.ceil(len(ds_ma) / so_cot)
    khung_tw = round(rong_khung * 567)                        # cạnh khung vuông (twips, 1 cm = 567)
    khe_tw = round((rong_cot - rong_khung) / 2 * 567)         # khe mỗi phía giữa khung và đường cắt
    bang = doc.add_table(rows=3 * so_hang, cols=3 * so_cot)
    bang.alignment = WD_TABLE_ALIGNMENT.CENTER
    bang.autofit = False                                   # bố cục cố định
    tbl_pr = bang._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str((khung_tw + 2 * khe_tw) * so_cot))
    layout = tbl_pr.find(qn("w:tblLayout"))
    le_o = OxmlElement("w:tblCellMar")                    # lề ô = 0 để kích thước đúng bằng số đã tính
    for canh in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{canh}")
        e.set(qn("w:w"), "0"); e.set(qn("w:type"), "dxa")
        le_o.append(e)
    layout.addnext(le_o)

    for c, cot in enumerate(bang.columns):
        cot.width = Twips(khung_tw if c % 3 == 1 else khe_tw)

    def dat_o(o, vien, le_trai_phai=0):
        """vien: {cạnh: (kiểu, độ dày, màu)}; đặt viền + lề trái/phải cho ô."""
        tc_pr = o._tc.get_or_add_tcPr()
        tc_w = tc_pr.find(qn("w:tcW"))
        vb = OxmlElement("w:tcBorders")
        for canh in ("top", "left", "bottom", "right"):
            kieu, day, mau = vien.get(canh, ("nil", "0", "auto"))
            e = OxmlElement(f"w:{canh}")
            for k, v in (("val", kieu), ("sz", day), ("space", "0"), ("color", mau)):
                e.set(qn(f"w:{k}"), v)
            vb.append(e)
        tc_w.addnext(vb)
        if le_trai_phai:
            mar = OxmlElement("w:tcMar")
            for canh in ("left", "right"):
                e = OxmlElement(f"w:{canh}")
                e.set(qn("w:w"), str(le_trai_phai)); e.set(qn("w:type"), "dxa")
                mar.append(e)
            vb.addnext(mar)
        o.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    NET_DUT, NET_DEN = ("dashed", "4", "A6A6A6"), ("single", "6", "000000")
    for r, hang in enumerate(bang.rows):
        tr = hang._tr.get_or_add_trPr()
        tr.append(OxmlElement("w:cantSplit"))              # không cắt đôi 1 hàng giữa 2 trang
        cao = OxmlElement("w:trHeight")                    # chiều cao CỐ ĐỊNH: hàng giữa = cạnh khung -> khung vuông
        cao.set(qn("w:val"), str(khung_tw if r % 3 == 1 else khe_tw)); cao.set(qn("w:hRule"), "exact")
        tr.append(cao)
        for c, o in enumerate(hang.cells):
            a_, b_ = r % 3, c % 3                          # vị trí trong khối 3x3 của 1 tem
            o.width = Twips(khung_tw if b_ == 1 else khe_tw)
            if a_ == 1 and b_ == 1:
                continue                                   # ô khung: xử lý ở dưới
            vien = {}
            if a_ == 0: vien["top"] = NET_DUT
            if a_ == 2: vien["bottom"] = NET_DUT
            if b_ == 0: vien["left"] = NET_DUT
            if b_ == 2: vien["right"] = NET_DUT
            dat_o(o, vien)
            p = o.paragraphs[0]                            # đoạn rỗng thật nhỏ để không làm phình hàng khe
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(1)
            p.add_run().font.size = Pt(1)

    for k, ma in enumerate(ds_ma):
        o_trong = bang.cell(3 * (k // so_cot) + 1, 3 * (k % so_cot) + 1)     # ô giữa của khối tem thứ k = khung vuông
        dat_o(o_trong, {canh: NET_DEN for canh in ("top", "left", "bottom", "right")}, le_trai_phai=85)
        p_anh = o_trong.paragraphs[0]
        p_anh.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_anh.paragraph_format.space_before = Pt(0)
        p_anh.paragraph_format.space_after = Pt(3)
        run_anh = p_anh.add_run()
        run_anh.font.size = Pt(1)                          # chữ nhỏ -> không có khoảng thừa dưới ảnh
        run_anh.add_picture(io.BytesIO(tao_anh_qr_cat_png(ma["noi_dung"])),
                            width=Cm(kich_thuoc_qr * so_module_qr(ma["noi_dung"]) / (so_module_qr(ma["noi_dung"]) + 8)))
        for inl in run_anh._r.findall(".//" + qn("wp:inline")):   # khoảng cách quanh ảnh = 0 (thiếu thì LibreOffice tự thêm lề làm ảnh lệch)
            for a_ in ("distT", "distB", "distL", "distR"):
                inl.set(a_, "0")
        p_chu = o_trong.add_paragraph()
        p_chu.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_chu.paragraph_format.space_before = p_chu.paragraph_format.space_after = Pt(0)
        r = p_chu.add_run(ma["noi_dung"])
        r.bold, r.font.size = True, Pt(co_chu)

    zoom = doc.settings.element.find(qn("w:zoom"))        # mẫu mặc định của python-docx thiếu thuộc tính này
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")

    bo_dem = io.BytesIO()
    doc.save(bo_dem)
    return bo_dem.getvalue()


def _tab_qr_san_pham():
    """Tab 1 của Tạo mã QR: tem QR sản phẩm (Kệ-Tên-Loại-STT, VD: K1-VFD-3P-01)."""

    nhan_muc("Danh sách sản phẩm")
    bang = st.data_editor(
        pd.DataFrame([{"Kệ": 1, "Tên sản phẩm": "", "Mã sản phẩm": "", "Tên loại sản phẩm": "", "Mã loại sản phẩm": "", "STT bắt đầu": 1, "Số lượng": 10}]),
        num_rows="dynamic", use_container_width=True, hide_index=True, key="qr_bang",
        column_config={
            "Kệ": st.column_config.NumberColumn("Kệ (số)", min_value=1, step=1, default=1, required=True),
            "Tên sản phẩm": st.column_config.TextColumn("Tên sản phẩm", help="Hiển thị trong ô / lịch sử khi quét nhập hàng. Bỏ trống thì hiển thị mã sản phẩm."),
            "Mã sản phẩm": st.column_config.TextColumn("Mã sản phẩm", required=True, help="Nằm trong mã QR, VD: VFD → K1-VFD-3P-01"),
            "Tên loại sản phẩm": st.column_config.TextColumn("Tên loại sản phẩm", help="Hiển thị trong ô / lịch sử khi quét nhập hàng. Bỏ trống thì hiển thị mã loại."),
            "Mã loại sản phẩm": st.column_config.TextColumn("Mã loại sản phẩm", required=True, help="Nằm trong mã QR, VD: 3P → K1-VFD-3P-01"),
            "STT bắt đầu": st.column_config.NumberColumn("STT bắt đầu", min_value=1, step=1, default=1),
            "Số lượng": st.column_config.NumberColumn("Số lượng", min_value=1, step=1, default=10),
        })
    nhan_muc("Tùy chọn in")
    kich_thuoc = st.number_input("Kích thước mỗi mã QR khi in (cm)", min_value=1.5, max_value=6.0, value=2.0,
                                 step=0.5, key="qr_kich_thuoc")

    if st.button("Tạo mã QR & file Word", type="primary", use_container_width=True, key="qr_tao"):
        ds, loi = kiem_tra_bang_qr(bang.to_dict("records"))
        if loi:
            st.session_state.pop("qr_ket_qua", None)
            for l in loi:
                st.error(l)
        else:
            try:
                with st.spinner(f"Đang tạo {len(ds)} mã QR..."):
                    docx_bytes = tao_file_word_qr(ds, kich_thuoc)
                    anh_xem_truoc = [tao_anh_qr_png(m["noi_dung"]) for m in ds[:8]]
            except ImportError:
                st.error("Chưa cài thư viện tạo file Word. Hãy chạy: `pip install python-docx` rồi khởi động lại app.")
            except RuntimeError as e:
                st.error(str(e))
            else:
                luu_ten_loai_day_du({(m["ten"], m["loai"], m["ten_loai_day_du"]) for m in ds})
                luu_ten_day_du({(m["ten"], m["ten_day_du"]) for m in ds})      # nhớ tên sản phẩm để quét nhập hàng hiển thị đúng
                da_co = [m["noi_dung"] for m in ds if tim_san_pham_trong_kho(m["ten"], m["loai"], m["stt"])]
                st.session_state["qr_ket_qua"] = {
                    "docx": docx_bytes, "so_ma": len(ds), "da_co": da_co, "anh": anh_xem_truoc,
                    "nhan": [m["noi_dung"] for m in ds[:8]], "ds": [m["noi_dung"] for m in ds],
                    "ten_file": f"ma_qr_san_pham_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                }

    kq = st.session_state.get("qr_ket_qua")
    if kq:
        st.divider()
        st.success(f"Đã tạo **{kq['so_ma']}** mã QR. Tải file Word bên dưới rồi in ra dán lên sản phẩm.")
        if kq["da_co"]:
            st.warning(f"Có {len(kq['da_co'])} mã **đang có sẵn trong kho** (nhập hàng sẽ bị chặn trùng): "
                       + ", ".join(f"`{m}`" for m in kq["da_co"][:10]) + (" ..." if len(kq["da_co"]) > 10 else ""))
        st.download_button(
            "TẢI FILE WORD (.docx)", data=kq["docx"], file_name=kq["ten_file"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", use_container_width=True, key="qr_tai")
        st.markdown(f"**Xem trước** (tối đa 8 mã đầu tiên):")
        st.image(kq["anh"], caption=kq["nhan"], width=130)
        with st.expander(f"Danh sách {kq['so_ma']} mã"):
            st.dataframe(pd.DataFrame({"Nội dung mã QR": kq["ds"]}), use_container_width=True, hide_index=True)


# ============ TẠO MÃ QR — TAB 2: QR VỊ TRÍ MỖI Ô (theo các ô đang có trong kệ) ============
def danh_sach_ma_o(cac_ke):
    """Mã QR của mọi ô trong các kệ đã chọn (theo thứ tự kệ -> hàng -> cột), đúng định dạng Kx Ty zz mà app quét được."""
    ds = []
    for ten_ke in cac_ke:
        ke = st.session_state.kho[ten_ke]
        for h in range(len(ke)):
            for c in range(len(ke[0])):
                ds.append({"ten_ke": ten_ke, "noi_dung": ma_o(ten_ke, h, c)})
    return ds


def _chon_tat_ca_ke(gia_tri):
    """Tích / bỏ tích tất cả kệ trong danh sách chọn kệ cần in QR."""
    for k in st.session_state.kho:
        st.session_state[f"qro_chon_{k}"] = gia_tri


def _tab_qr_vi_tri_o():
    """Tab 2 của Tạo mã QR: tem QR vị trí từng ô, lấy theo các kệ đang có."""
    if not co_ke():
        return

    nhan_muc("Chọn kệ cần in")
    for k in st.session_state.kho:
        st.session_state.setdefault(f"qro_chon_{k}", True)          # mặc định chọn hết
    c_het, c_bo = st.columns(2)
    c_het.button("Chọn tất cả", on_click=_chon_tat_ca_ke, args=(True,), use_container_width=True, key="qro_chon_het")
    c_bo.button("Bỏ chọn tất cả", on_click=_chon_tat_ca_ke, args=(False,), use_container_width=True, key="qro_bo_het")
    khung = st.container(height=280, border=True) if len(st.session_state.kho) > 6 else st.container(border=True)
    with khung:
        for k, ke in st.session_state.kho.items():
            st.checkbox(f"**{k}** — {len(ke)} hàng × {len(ke[0])} cột = {len(ke) * len(ke[0])} ô",
                        key=f"qro_chon_{k}")
    cac_ke = [k for k in st.session_state.kho if st.session_state.get(f"qro_chon_{k}")]
    ds = danh_sach_ma_o(cac_ke)
    if cac_ke:
        pass
    else:
        st.warning("Chưa chọn kệ nào. Hãy tích chọn ít nhất 1 kệ trong danh sách.")

    nhan_muc("Tùy chọn in")
    kich_thuoc = st.number_input("Kích thước mỗi mã QR khi in (cm)", min_value=1.5, max_value=6.0, value=2.0,
                                 step=0.5, key="qro_kich_thuoc")

    if st.button("Tạo QR ô & file Word", type="primary", use_container_width=True, key="qro_tao"):
        if not ds:
            st.session_state.pop("qro_ket_qua", None)
            st.error("Hãy chọn ít nhất 1 kệ.")
        elif len(ds) > MAX_MA_QR:
            st.session_state.pop("qro_ket_qua", None)
            st.error(f"Tổng cộng {len(ds)} mã, vượt giới hạn {MAX_MA_QR} mã mỗi lần. Hãy chọn ít kệ hơn.")
        else:
            try:
                with st.spinner(f"Đang tạo {len(ds)} mã QR..."):
                    docx_bytes = tao_file_word_qr(ds, kich_thuoc, tieu_de="Tem QR vị trí ô")
                    anh_xem_truoc = [tao_anh_qr_png(m["noi_dung"]) for m in ds[:8]]
            except ImportError:
                st.error("Chưa cài thư viện tạo file Word. Hãy chạy: `pip install python-docx` rồi khởi động lại app.")
            except RuntimeError as e:
                st.error(str(e))
            else:
                st.session_state["qro_ket_qua"] = {
                    "docx": docx_bytes, "so_ma": len(ds), "anh": anh_xem_truoc,
                    "nhan": [m["noi_dung"] for m in ds[:8]], "ds": [m["noi_dung"] for m in ds],
                    "ten_file": f"qr_vi_tri_o_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                }

    kq = st.session_state.get("qro_ket_qua")
    if kq:
        st.divider()
        st.success(f"Đã tạo **{kq['so_ma']}** mã QR vị trí ô. Tải file Word bên dưới rồi in ra dán lên từng ô.")
        st.download_button(
            "TẢI FILE WORD (.docx)", data=kq["docx"], file_name=kq["ten_file"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", use_container_width=True, key="qro_tai")
        st.markdown("**Xem trước** (tối đa 8 mã đầu tiên):")
        st.image(kq["anh"], caption=kq["nhan"], width=130)
        with st.expander(f"Danh sách {kq['so_ma']} mã"):
            st.dataframe(pd.DataFrame({"Nội dung mã QR": kq["ds"]}), use_container_width=True, hide_index=True)


def man_hinh_tao_qr():
    tieu_de_trang("Tạo mã QR")
    tab_sp, tab_o = st.tabs(["QR sản phẩm", "QR vị trí ô"])
    with tab_sp:
        _tab_qr_san_pham()
    with tab_o:
        _tab_qr_vi_tri_o()
    nut_quay_lai()


# ============ CHỨC NĂNG: TẠO KỆ HÀNG ============
def man_hinh_tao_ke():
    """Nhập số hàng (tầng) và số cột (ô mỗi tầng) để tạo 1 kệ mới."""
    tieu_de_trang("Quản lý kệ hàng")

    st.session_state.pop("tao_ke_tb", None)

    nhan_muc("Kệ hiện có")
    if co_ke():
        dong = []
        for ten_ke, ke in st.session_state.kho.items():
            so_hang, so_cot = len(ke), len(ke[0])
            dang_dung = sum(1 for hang in ke for o in hang for vt in o if vt is not None)
            so_vt = len(ke[0][0])
            dong.append({"Kệ": ten_ke, "Số hàng": so_hang, "Số cột": so_cot, "Vị trí mỗi ô": so_vt,
                         "Tổng số ô": so_hang * so_cot,
                         "Vị trí đang dùng": f"{dang_dung}/{so_hang * so_cot * so_vt}"})
        st.dataframe(dong, use_container_width=True, hide_index=True)
    else:
        pass

    nhan_muc("Tạo kệ mới")
    c1, c2, c3 = st.columns(3)
    so_hang = int(c1.number_input("Số hàng (số tầng)", min_value=1, max_value=MAX_HANG, value=5, step=1, key="tao_ke_hang"))
    so_cot = int(c2.number_input("Số cột (số ô mỗi tầng)", min_value=1, max_value=MAX_COT, value=4, step=1, key="tao_ke_cot"))
    so_vt = int(c3.number_input("Số vị trí mỗi ô", min_value=1, max_value=MAX_VI_TRI, value=SO_VI_TRI, step=1, key="tao_ke_vt"))
    n = so_ke_tiep_theo()
    if st.button("Tạo kệ", type="primary", use_container_width=True, key="tao_ke_ok"):
        ten_ke = them_ke(so_hang, so_cot, so_vt)
        st.session_state["tao_ke_tb"] = (f"Đã tạo **{ten_ke}** ({so_hang} hàng × {so_cot} cột, mỗi ô {so_vt} vị trí). "
                                         f"Bây giờ có thể nhập / xuất hàng.")
        st.rerun()

    if co_ke():
        nhan_muc("Sửa kệ")
        ten_sua = st.selectbox("Chọn kệ cần sửa", list(st.session_state.kho), key="sua_ke_chon")
        ke_sua = st.session_state.kho[ten_sua]
        h0, c0, v0 = len(ke_sua), len(ke_sua[0]), len(ke_sua[0][0])
        s1, s2, s3 = st.columns(3)
        h1 = int(s1.number_input("Số hàng (số tầng)", min_value=1, max_value=max(MAX_HANG, h0), value=h0, step=1, key=f"sua_hang_{ten_sua}"))
        c1_ = int(s2.number_input("Số cột (số ô mỗi tầng)", min_value=1, max_value=max(MAX_COT, c0), value=c0, step=1, key=f"sua_cot_{ten_sua}"))
        v1 = int(s3.number_input("Số vị trí mỗi ô", min_value=1, max_value=max(MAX_VI_TRI, v0), value=v0, step=1, key=f"sua_vt_{ten_sua}"))
        if (h1, c1_, v1) == (h0, c0, v0):
            co_the_sua = False
        else:
            ly_do = kiem_tra_sua_ke(ten_sua, h1, c1_, v1)
            if ly_do:
                for ld in ly_do:
                    st.warning(ld + " Hãy xuất hàng ở đó trước.")
                co_the_sua = False
            else:
                co_the_sua = True
        if st.button("Lưu thay đổi", type="primary", use_container_width=True, disabled=not co_the_sua, key="sua_ke_btn"):
            ok, loi = sua_ke(ten_sua, h1, c1_, v1)
            if ok:
                st.session_state["tao_ke_tb"] = f"Đã sửa **{ten_sua}**: {h1} hàng × {c1_} cột × {v1} vị trí mỗi ô."
                st.rerun()
            else:
                st.error(loi)

    if co_ke():
        nhan_muc("Xóa kệ")
        ten_xoa = st.selectbox("Chọn kệ cần xóa", list(st.session_state.kho), key="xoa_ke_chon")
        dang_dung = ke_dang_dung(ten_xoa)
        if dang_dung > 0:
            with st.container(key="xoa_ke_loi"):
                st.error(f"**{ten_xoa}** đang có sản phẩm không thể xóa.")
        else:
            xac_nhan = st.checkbox(f"Tôi chắc chắn muốn xóa {ten_xoa}", key=f"xoa_ke_ok_{ten_xoa}")
            if st.button(f"Xóa {ten_xoa}", use_container_width=True, disabled=not xac_nhan, key="xoa_ke_btn"):
                if xoa_ke(ten_xoa):
                    st.session_state["tao_ke_tb"] = f"Đã xóa **{ten_xoa}**."
                    st.rerun()

    nut_quay_lai()


def man_hinh_tao_ke_dau_tien():
    """BƯỚC 2 của quy trình: Đăng nhập -> TẠO KỆ -> Trang chính.
    Tài khoản chưa có kệ nào thì CHỈ thấy màn hình này, tạo kệ xong mới vào được trang chính."""
    _, giua, _ = st.columns([1, 2, 1])
    with giua:
        tieu_de_trang("Tạo kệ hàng đầu tiên")
        st.session_state.pop("tao_ke_tb", None)
        c1, c2, c3 = st.columns(3)
        so_hang = int(c1.number_input("Số hàng (số tầng)", min_value=1, max_value=MAX_HANG, value=5, step=1, key="bb_hang"))
        so_cot = int(c2.number_input("Số cột (số ô mỗi tầng)", min_value=1, max_value=MAX_COT, value=4, step=1, key="bb_cot"))
        so_vt = int(c3.number_input("Số vị trí mỗi ô", min_value=1, max_value=MAX_VI_TRI, value=SO_VI_TRI, step=1, key="bb_vt"))
        n = so_ke_tiep_theo()
        if st.button("Tạo kệ và vào trang chính", type="primary", use_container_width=True, key="bb_tao"):
            ten_ke = them_ke(so_hang, so_cot, so_vt)
            st.session_state["tb_trang_chu"] = (f"Đã tạo **{ten_ke}** ({so_hang} hàng × {so_cot} cột, mỗi ô {so_vt} vị trí). "
                                                f"Chào mừng bạn vào trang chính!")
            st.session_state.chuc_nang = None
            st.rerun()


# ============ ĐĂNG NHẬP / TẠO TÀI KHOẢN ============
def bam_mat_khau(mat_khau, muoi_hex):
    """Băm mật khẩu bằng PBKDF2-SHA256 (200.000 vòng) với muối riêng của từng tài khoản."""
    return hashlib.pbkdf2_hmac("sha256", mat_khau.encode("utf-8"), bytes.fromhex(muoi_hex), 200_000).hex()


def tao_tai_khoan(ten, mat_khau, nhap_lai):
    """Tạo tài khoản mới. Trả về (True, tên_hiển_thị) hoặc (False, thông báo lỗi)."""
    ten = (ten or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9_.\-]{3,30}", ten):
        return False, "Tên tài khoản dài 3–30 ký tự, chỉ gồm chữ không dấu, số, dấu gạch dưới (_), gạch ngang (-) hoặc dấu chấm (.)."
    if len(mat_khau or "") < MK_TOI_THIEU:
        return False, f"Mật khẩu phải có ít nhất {MK_TOI_THIEU} ký tự."
    if mat_khau != nhap_lai:
        return False, "Hai lần nhập mật khẩu không giống nhau."
    muoi = secrets.token_hex(16)
    if not them_tai_khoan(ten, muoi, bam_mat_khau(mat_khau, muoi)):    # phân biệt hoa/thường: chỉ trùng khi giống từng chữ
        return False, f"Tên tài khoản **{ten}** đã có người dùng. Hãy chọn tên khác."
    return True, ten


def kiem_tra_dang_nhap(ten, mat_khau):
    """Trả về tên hiển thị nếu đúng tài khoản + mật khẩu, ngược lại None. Luôn tốn thời gian băm như nhau."""
    tk = lay_tai_khoan((ten or "").strip())      # phải gõ ĐÚNG chữ hoa / chữ thường
    muoi = tk["muoi"] if tk else "00" * 16
    bam = bam_mat_khau(mat_khau or "", muoi)
    if tk and hmac.compare_digest(bam, tk["mat_khau"]):
        return tk["ten"]
    return None


def _dang_nhap_thanh_cong(ten_hien_thi):
    st.session_state.clear()                  # xóa sạch dữ liệu phiên của người trước
    st.session_state["nguoi_dung"] = ten_hien_thi
    st.session_state["chuc_nang"] = None
    st.session_state["kho"] = doc_du_lieu(ten_hien_thi)
    st.session_state["_dat_cookie"] = tao_phien(ten_hien_thi)      # lát nữa đặt cookie để F5 không bị văng
    st.rerun()


def man_hinh_dang_nhap():
    st.markdown(_gon_html("<style>.st-key-dang_xuat button{ min-height:40px !important; }</style>"), unsafe_allow_html=True)
    _, giua, _ = st.columns([1, 2, 1])
    with giua:
        tieu_de_trang("Quản lý kho hàng")
        tab_dn, tab_tao = st.tabs(["Đăng nhập", "Tạo tài khoản"])

        with tab_dn:
            with st.form("form_dang_nhap"):
                ten = st.text_input("Tên tài khoản", key="dn_ten")
                mk = st.text_input("Mật khẩu", type="password", key="dn_mk")
                gui = st.form_submit_button("Đăng nhập", type="primary", use_container_width=True)
            if gui:
                con_lai = st.session_state.get("dn_khoa_den", 0) - time.time()
                if con_lai > 0:
                    st.error(f"Nhập sai quá nhiều lần. Vui lòng thử lại sau {int(con_lai) + 1} giây.")
                elif not ten.strip() or not mk:
                    st.error("Vui lòng nhập tên tài khoản và mật khẩu.")
                else:
                    ten_hien_thi = kiem_tra_dang_nhap(ten, mk)
                    if ten_hien_thi:
                        _dang_nhap_thanh_cong(ten_hien_thi)
                    else:
                        so_sai = st.session_state.get("dn_sai", 0) + 1
                        st.session_state["dn_sai"] = so_sai
                        if so_sai >= SAI_TOI_DA:
                            st.session_state["dn_khoa_den"] = time.time() + KHOA_GIAY
                            st.session_state["dn_sai"] = 0
                            st.error(f"Nhập sai {SAI_TOI_DA} lần. Đăng nhập bị tạm khóa {KHOA_GIAY} giây.")
                        else:
                            st.error("Sai tên tài khoản hoặc mật khẩu.")

        with tab_tao:
            with st.form("form_tao_tai_khoan"):
                ten_moi = st.text_input("Tên tài khoản mới", key="tk_ten", help="3–30 ký tự: chữ không dấu, số, _ - . (phân biệt chữ hoa / chữ thường)")
                mk_moi = st.text_input(f"Mật khẩu (tối thiểu {MK_TOI_THIEU} ký tự)", type="password", key="tk_mk")
                mk_lai = st.text_input("Nhập lại mật khẩu", type="password", key="tk_mk2")
                tao = st.form_submit_button("Tạo tài khoản", type="primary", use_container_width=True)
            if tao:
                ok, ket_qua = tao_tai_khoan(ten_moi, mk_moi, mk_lai)
                if ok:
                    _dang_nhap_thanh_cong(ket_qua)        # tạo xong tự đăng nhập luôn
                else:
                    st.error(ket_qua)


def thanh_nguoi_dung(ten_nguoi_dung):
    """Dòng nhỏ trên cùng mỗi trang: đang đăng nhập bằng tài khoản nào + nút đăng xuất."""
    st.markdown(_gon_html("<style>.st-key-dang_xuat button{ min-height:40px !important; }</style>"), unsafe_allow_html=True)
    c1, c2 = st.columns([4, 1])
    c1.caption(f"Đang đăng nhập: **{ten_nguoi_dung}**")
    if c2.button("Đăng xuất", key="dang_xuat", use_container_width=True):
        huy_phien(_cookie_hien_tai() or st.session_state.get("_ma_phien"))   # hủy phiên trên server
        st.session_state.clear()      # xóa hết: người tiếp theo không thấy kho / QR của người này
        st.session_state["_xoa_cookie"] = True
        st.rerun()


# ============ ĐIỀU HƯỚNG CHÍNH ============
st.markdown(_gon_html(CSS_CHUNG), unsafe_allow_html=True)      # giao diện dùng chung cho mọi trang
nguoi_dung = st.session_state.get("nguoi_dung")      # đọc 1 lần, dùng lại bên dưới (không đọc lại từ session)
if not nguoi_dung and not st.session_state.get("_xoa_cookie"):
    # Tải lại trang -> session_state rỗng. Thử khôi phục đăng nhập từ cookie.
    ten_khoi_phuc = tim_phien(_cookie_hien_tai())
    if ten_khoi_phuc:
        try:
            st.session_state["nguoi_dung"] = ten_khoi_phuc
            st.session_state["chuc_nang"] = None
            st.session_state["kho"] = doc_du_lieu(ten_khoi_phuc)
            st.session_state["_ma_phien"] = _cookie_hien_tai()
            nguoi_dung = ten_khoi_phuc
        except Exception:
            st.session_state.clear()
            nguoi_dung = None
if not nguoi_dung:
    if st.session_state.pop("_xoa_cookie", False):
        _js_cookie(None)              # vừa đăng xuất -> xóa cookie trên trình duyệt
    man_hinh_dang_nhap()              # chưa đăng nhập -> chỉ thấy màn hình đăng nhập / tạo tài khoản
    st.stop()
ma_moi = st.session_state.pop("_dat_cookie", None)
if ma_moi:
    st.session_state["_ma_phien"] = ma_moi
    _js_cookie(ma_moi)                # vừa đăng nhập -> ghi cookie để lần tải lại sau vẫn còn đăng nhập
if "kho" not in st.session_state:
    # Phiên có tên người dùng nhưng thiếu kho (VD: app vừa được sửa/nạp lại giữa chừng) -> nạp lại từ file thay vì sập.
    try:
        st.session_state["kho"] = doc_du_lieu(nguoi_dung)
        st.session_state.setdefault("chuc_nang", None)
    except Exception:
        st.session_state.clear()      # tài khoản không còn / dữ liệu hỏng -> về màn hình đăng nhập
        st.rerun()
thanh_nguoi_dung(nguoi_dung)
if not co_ke():
    man_hinh_tao_ke_dau_tien()     # BƯỚC 2: chưa có kệ nào -> bắt buộc tạo kệ, chưa cho vào trang chính / chức năng nào
elif st.session_state.chuc_nang is None:
    st.session_state.pop("tb_trang_chu", None)
    trang_chu()
elif st.session_state.chuc_nang == "kiem_tra":
    man_hinh_kiem_tra()
elif st.session_state.chuc_nang == "quet_kiem_tra":
    man_hinh_quet_kiem_tra()
elif st.session_state.chuc_nang == "nhap":
    man_hinh_nhap()
elif st.session_state.chuc_nang == "quet_nhap":
    man_hinh_quet_nhap()
elif st.session_state.chuc_nang == "xuat":
    man_hinh_xuat()
elif st.session_state.chuc_nang == "quet_xuat":
    man_hinh_quet_xuat()
elif st.session_state.chuc_nang == "lich_su":
    man_hinh_lich_su()
elif st.session_state.chuc_nang == "tao_qr":
    man_hinh_tao_qr()
elif st.session_state.chuc_nang == "tao_ke":
    man_hinh_tao_ke()