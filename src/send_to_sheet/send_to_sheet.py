import requests
import json
import time
import os

# Địa chỉ API vMix Web Controller (mặc định)
VMIX_API_URL = "http://127.0.0.1:8088/api/"

# Lấy Tran.Text từ Input 1 của vMix
import xml.etree.ElementTree as ET
def get_tran_from_vmix():
    try:
        resp = requests.get(VMIX_API_URL)
        if resp.status_code == 200:
            root = ET.fromstring(resp.text)
            for inp in root.findall("input"):
                if inp.get("number") == "1":
                    # Ưu tiên tìm <text name="Tran.Text">
                    for text in inp.findall("text"):
                        if text.get("name") == "Tran.Text":
                            return text.text
                    # Nếu không có, thử <text name="Tran">
                    for text in inp.findall("text"):
                        if text.get("name") == "Tran":
                            return text.text
                    # Nếu không có, thử lấy thuộc tính trực tiếp
                    if "Tran.Text" in inp.attrib:
                        return inp.get("Tran.Text")
                    if "Tran" in inp.attrib:
                        return inp.get("Tran")
        else:
            print("Không kết nối được vMix API:", resp.status_code)
    except Exception as e:
        print("Lỗi lấy Tran từ vMix:", e)
    return None

# URL Apps Script (deploy ID của bạn)
URL = "https://script.google.com/macros/s/AKfycbwPZynhHVwOwTttkZ3-WR7KAbRD1zexdgfaU2T0BWSfF2FXpsyvPV_UxdFficrV0TWATw/exec"

# File do vMix ghi
DATA_FILE = r"C:\Users\Public\Documents\vmix_data.txt"

def send_to_sheet(data):
    try:
        resp = requests.get(URL, params=data)
        print("Đã gửi:", data, "→ Response:", resp.text)
    except Exception as e:
        print("Lỗi gửi:", e)

def main():
    print("Bắt đầu theo dõi file vMix...")
    last_mtime = None

    while True:
        try:
            if os.path.exists(DATA_FILE):
                mtime = os.path.getmtime(DATA_FILE)  # thời gian chỉnh sửa file
                if last_mtime is None or mtime != last_mtime:
                    last_mtime = mtime
                    with open(DATA_FILE, "r", encoding="utf-8-sig") as f:  # fix BOM
                        content = f.read().strip()
                    if content:
                        try:
                            data = json.loads(content)

                            # chuẩn hóa AvgA/B: đổi dấu phẩy thành dấu chấm
                            for key in ["AvgA", "AvgB"]:
                                if key in data:
                                    data[key] = data[key].replace(",", ".")

                            # Lấy Tran từ vMix và thêm vào data
                            tran_value = get_tran_from_vmix()
                            if tran_value:
                                data["Tran"] = tran_value
                            else:
                                print("Không lấy được Tran từ vMix, sẽ không gửi trường Tran!")
                            send_to_sheet(data)
                        except json.JSONDecodeError:
                            print("File không phải JSON hợp lệ:", content)
            else:
                print("Chưa thấy file:", DATA_FILE)

        except Exception as e:
            print("Lỗi:", e)

        time.sleep(1)  # kiểm tra mỗi giây

if __name__ == "__main__":
    main()
