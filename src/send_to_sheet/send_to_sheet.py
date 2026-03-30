import requests
import json
import time
import os
import xml.etree.ElementTree as ET

# ─── Cấu hình ─────────────────────────────────────────────────────────────────
VMIX_API_URL = "http://127.0.0.1:8088/api/"
WEB_BASE_URL = "https://hbsf.com.vn"       # ← Địa chỉ web server

# URL Apps Script Google Sheets (giữ nguyên logic cũ)
GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbwPZynhHVwOwTttkZ3-WR7KAbRD1zexdgfaU2T0BWSfF2FXpsyvPV_UxdFficrV0TWATw/exec"

DATA_FILE = r"C:\Users\Public\Documents\vmix_data.txt"
# ──────────────────────────────────────────────────────────────────────────────


def get_tran_from_vmix():
    """Lấy giá trị Tran.Text từ Input 1 của vMix."""
    try:
        resp = requests.get(VMIX_API_URL, timeout=3)
        if resp.status_code == 200:
            root = ET.fromstring(resp.text)
            for inp in root.findall("input"):
                if inp.get("number") == "1":
                    for text in inp.findall("text"):
                        if text.get("name") == "Tran.Text":
                            return text.text
                    for text in inp.findall("text"):
                        if text.get("name") == "Tran":
                            return text.text
                    if "Tran.Text" in inp.attrib:
                        return inp.get("Tran.Text")
                    if "Tran" in inp.attrib:
                        return inp.get("Tran")
        else:
            print("Không kết nối được vMix API:", resp.status_code)
    except Exception as e:
        print("Lỗi lấy Tran từ vMix:", e)
    return None


def send_to_sheet(data):
    """Gửi dữ liệu lên Google Sheets qua Apps Script (logic cũ)."""
    try:
        resp = requests.get(GOOGLE_SHEET_URL, params=data)
        print("Đã gửi Google Sheets:", data, "→ Response:", resp.text)
    except Exception as e:
        print("Lỗi gửi Google Sheets:", e)


def finalize_match(tran_value):
    """
    Gọi API kết thúc trận đấu trên web.
    Backend tự xác định vòng loại hay vòng chính dựa theo ngày hôm nay,
    tính lại avg và cập nhật winner_id.
    """
    url = f"{WEB_BASE_URL.rstrip('/')}/api/matches/finalize-by-tran"
    try:
        resp = requests.post(url, json={"tran": int(tran_value)}, timeout=10)
        if resp.status_code == 200:
            print(f"✅ Kết thúc trận {tran_value} thành công:", resp.json().get("message", ""))
        else:
            try:
                err = resp.json().get("message", resp.text[:300])
            except Exception:
                err = resp.text[:300]
            print(f"❌ Lỗi kết thúc trận {tran_value}: HTTP {resp.status_code} — {err}")
    except Exception as e:
        print(f"❌ Lỗi kết nối server khi kết thúc trận {tran_value}:", e)


def main():
    print("Bắt đầu theo dõi file vMix...")
    last_mtime = None

    while True:
        try:
            if os.path.exists(DATA_FILE):
                mtime = os.path.getmtime(DATA_FILE)
                if last_mtime is None or mtime != last_mtime:
                    last_mtime = mtime
                    with open(DATA_FILE, "r", encoding="utf-8-sig") as f:  # fix BOM
                        content = f.read().strip()
                    if content:
                        try:
                            data = json.loads(content)
                        except json.JSONDecodeError:
                            print("File không phải JSON hợp lệ:", content)
                            time.sleep(1)
                            continue

                        # chuẩn hóa AvgA/B: đổi dấu phẩy thành dấu chấm
                        for key in ["AvgA", "AvgB"]:
                            if key in data:
                                data[key] = data[key].replace(",", ".")

                        # Lấy Tran từ vMix
                        tran_value = get_tran_from_vmix()
                        if tran_value:
                            data["Tran"] = tran_value
                        else:
                            print("Không lấy được Tran từ vMix, sẽ không gửi trường Tran!")

                        # 1. Gửi lên Google Sheets (logic cũ, giữ nguyên)
                        send_to_sheet(data)

                        # 2. Gọi API web kết thúc trận đấu (logic mới)
                        if tran_value:
                            finalize_match(tran_value)
            else:
                print("Chưa thấy file:", DATA_FILE)

        except Exception as e:
            print("Lỗi:", e)

        time.sleep(1)  # kiểm tra mỗi giây


if __name__ == "__main__":
    main()
