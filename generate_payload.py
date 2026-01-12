from openpyxl import load_workbook
import json
import os
from datetime import datetime

# =====================================================
# BASE DIRECTORY (ตำแหน่งไฟล์ generate_payload.py)
# =====================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ================= COLUMN MAPPING =================
# อ้างอิงจาก Excel Row 22 (index เริ่มที่ 1)
COL = {
    "system_code": 1,
    "approval_flag": 2,
    "call_back_url": 3,

    "doc_no": 4,
    "doc_ref": 5,
    "amount": 6,
    "remark": 7,

    "group_code": 8,
    "account_code": 9,
    "disbursement_type": 10,

    "ga_no": 11,
    "payment_type": 12,
    "payee_type": 13,
    "payee_name": 14,

    "title_en": 15,
    "title_th": 16,
    "fname_en": 17,
    "fname_th": 18,
    "lname_en": 19,
    "lname_th": 20,

    "mobile_no": 21,
    "idcard_no": 22,
    "tax_id": 23,
    "passport_no": 24,
    "code": 25,

    "bank_code": 26,
    "branch_code": 27,
    "account_no": 28,
    "account_name": 29,
    "media_clearing_type_code": 30,

    "cheque_info": 31,
    "k_trade_info": 32,
    "relation": 33,
    "date_of_birth": 34,
    "pay_rate": 35,
    "payee_amount": 36,

    "delivery_address_1": 37,
    "delivery_address_2": 38,
    "mailing_zip_code": 39,

    "tax_rate": 40,
    "boi": 41,
    "boi_start_date": 42,
    "boi_expiry_date": 43,
}
# ==================================================


# ================= HELPER FUNCTIONS =================
def cell(row, key):
    """อ่านค่าจาก row ตาม column mapping"""
    val = row[COL[key] - 1]
    if val in ("", None, "null", "NULL"):
        return None
    return val


def as_str(val):
    """แปลงค่าเป็น string"""
    if val is None:
        return ""
    if isinstance(val, float):
        return str(int(val))
    return str(val)


def as_float(val):
    """แปลงค่าเป็น float"""
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def as_bool(val):
    """แปลงค่าเป็น boolean"""
    if isinstance(val, bool):
        return val
    if val in (1, "1", "Y", "y", "true", "TRUE"):
        return True
    return False


def as_date(val):
    """แปลงวันที่เป็น ISO 8601"""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.isoformat() + "Z"
    return str(val)
# ===================================================


# ================= MAIN FUNCTION =================
def generate_payload(
    excel_path=None,
    sheet_name="API_Doc",
    data_start_row=23,
    debug=False
):
    """
    อ่าน Excel แล้ว generate JSON payload
    return: list ของ path ไฟล์ที่สร้าง
    """

    # ---------- PATH CONFIG ----------
    if excel_path is None:
        excel_path = os.path.join(BASE_DIR, "API_Transaction.xlsx")

    output_dir = os.path.join(BASE_DIR, "payloads")
    os.makedirs(output_dir, exist_ok=True)
    # --------------------------------

    wb = load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name]

    generated_files = []
    running_no = 1

    for row in ws.iter_rows(min_row=data_start_row, values_only=True):
        if not any(row):
            continue

        # ---------- DEBUG FIRST ROW ----------
        if debug and running_no == 1:
            print("===== DEBUG FIRST DATA ROW =====")
            for k, c in COL.items():
                print(f"{k:25} -> col {c}: {row[c-1]}")
            print("================================\n")

        payload = {
            "system_code": as_str(cell(row, "system_code")),
            "approval_flag": as_bool(cell(row, "approval_flag")),
            "call_back_url": as_str(cell(row, "call_back_url")),

            "approval": {
                "doc_no": as_str(cell(row, "doc_no")),
                "doc_ref": as_str(cell(row, "doc_ref")),
                "amount": as_float(cell(row, "payee_amount")),
                "remark": as_str(cell(row, "remark")),

                "doc_meta_data": {
                    "group_code": as_str(cell(row, "group_code")),
                    "account_code": as_str(cell(row, "account_code")),
                    "disbursement_type": as_str(cell(row, "disbursement_type")),
                },

                "payees": {
                    "ga_no": as_str(cell(row, "ga_no")),
                    "payee_info": [
                        {
                            "payment_type": as_str(cell(row, "payment_type")),
                            "payee_type": as_str(cell(row, "payee_type")),
                            "payee_name": as_str(cell(row, "payee_name")),

                            "title": {
                                "en_US": as_str(cell(row, "title_en")),
                                "th_TH": as_str(cell(row, "title_th")),
                            },
                            "first_name": {
                                "en_US": as_str(cell(row, "fname_en")),
                                "th_TH": as_str(cell(row, "fname_th")),
                            },
                            "last_name": {
                                "en_US": as_str(cell(row, "lname_en")),
                                "th_TH": as_str(cell(row, "lname_th")),
                            },

                            "mobile_no": as_str(cell(row, "mobile_no")),
                            "idcard_no": as_str(cell(row, "idcard_no")),
                            "tax_id": as_str(cell(row, "tax_id")),
                            "passport_no": as_str(cell(row, "passport_no")),
                            "code": as_str(cell(row, "code")),

                            "bank_info": {
                                "bank_code": as_str(cell(row, "bank_code")),
                                "branch_code": as_str(cell(row, "branch_code")),
                                "account_no": as_str(cell(row, "account_no")),
                                "account_name": as_str(cell(row, "account_name")),
                                "media_clearing_type_code": as_str(cell(row, "media_clearing_type_code")),
                            },

                            "cheque_info": None,
                            "k_trade_info": None,
                            "relation": as_str(cell(row, "relation")),
                            "date_of_birth": as_date(cell(row, "date_of_birth")),
                            "pay_rate": as_float(cell(row, "pay_rate")),
                            "amount": as_float(cell(row, "payee_amount")),

                            "delivery_address_1": as_str(cell(row, "delivery_address_1")),
                            "delivery_address_2": as_str(cell(row, "delivery_address_2")),
                            "mailing_zip_code": as_str(cell(row, "mailing_zip_code")),

                            "tax": {
                                "tax_rate": as_float(cell(row, "tax_rate")),
                                "boi": as_bool(cell(row, "boi")),
                                "boi_start_date": as_date(cell(row, "boi_start_date")),
                                "boi_expiry_date": as_date(cell(row, "boi_expiry_date")),
                            },
                        }
                    ],
                    "committees": []
                }
            }
        }

        file_name = f"payload_{running_no:03}.json"
        file_path = os.path.join(output_dir, file_name)

        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        print(f"✅ Generated: {file_path}")
        generated_files.append(file_path)
        running_no += 1

    return generated_files
# ===================================================


# ================= SCRIPT ENTRY =================
if __name__ == "__main__":
    files = generate_payload(debug=True)
    print(f"\n🎉 Generated {len(files)} payload files")
# ===================================================
