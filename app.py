import streamlit as st
import os
import json
import zipfile
import requests
from generate_payload import generate_payload

# =====================================================
# CONFIG
# =====================================================
API_URL = "https://gisx-qa.muangthai.co.th/api/v1/disbursement/batches/v1/inbound/disbursements"

PAYLOAD_DIR = "payloads"
TEMPLATE_FILE = "API_Transaction.xlsx"

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="API Payload Generator",
    layout="centered"
)

st.title("🚀 Excel → JSON → API Disbursement")
st.caption("Generate payload & send to Disbursement API")

# =====================================================
# HOW TO USE (WIZARD)
# =====================================================
st.markdown("## 🧭 วิธีใช้งาน")

with st.expander("📘 Step 1 : เตรียมไฟล์ Excel", expanded=True):
    st.markdown("""
**1. ดาวน์โหลด Excel Template**
- กดปุ่ม **Download Excel Template**
- ไฟล์จะมี 2 Sheet:
  - `API_Doc_Agent_Broker`
  - `API_Doc_Company`

**2. โครงสร้างไฟล์**
- Header อยู่ที่ **Row 22**
- ข้อมูลเริ่มที่ **Row 23**
- ❌ ห้ามลบหรือสลับลำดับ Column

**3. การตั้งชื่อไฟล์ JSON**
- ใช้ข้อมูลจาก Column **A–D**
- รูปแบบชื่อไฟล์:
A_B_C_D.json 
**ตัวอย่าง** GPM_Agent_Bank_transfer_DT0001.json""")

with st.expander("🧩 Step 2 : Generate JSON Payload"):
    st.markdown("""
**1. Upload Excel**
- เลือกไฟล์ Excel ที่กรอกข้อมูลเรียบร้อยแล้ว

**2. Generate Payload**
- กดปุ่ม **Generate JSON Payload**
- ระบบจะ:
  - อ่านข้อมูลจากทุก Sheet
  - Generate JSON แยก **1 แถว = 1 ไฟล์**
  - เก็บไฟล์ไว้ในโฟลเดอร์ `payloads/`

**3. Download (ถ้าต้องการ)**
- สามารถดาวน์โหลด JSON ทั้งหมดเป็น ZIP ได้
""")

with st.expander("🚀 Step 3 : Select & Send to API"):
    st.markdown("""
**1. กรอกข้อมูล API**
- Bearer Token
- `x-user-name`

**2. เลือกไฟล์**
- ☑️ เลือกไฟล์ที่ต้องการยิง API
- ใช้ปุ่ม:
  - **Select All**
  - **Unselect All**

**3. ยิง API**
- กดปุ่ม **Send to API**
- ระบบจะ:
  - ยิงทีละไฟล์ (ตามลำดับ)
  - แสดงผลลัพธ์แยกต่อไฟล์
  - ไฟล์ที่ยิงสำเร็จจะถูก **disable checkbox**

**4. Result**
- แสดง:
  - HTTP Status
  - Response Body
  - Error (ถ้ามี)
- สรุปไฟล์ที่ Fail หลังยิงครบ
""")

st.divider()


# =====================================================
# DOWNLOAD TEMPLATE
# =====================================================
st.markdown("## 📥 Excel Template")

if os.path.exists(TEMPLATE_FILE):
    with open(TEMPLATE_FILE, "rb") as f:
        st.download_button(
            "⬇️ Download Excel Template (Current)",
            f,
            file_name=TEMPLATE_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.error("❌ ไม่พบไฟล์ Template")

st.divider()

# =====================================================
# UPLOAD EXCEL
# =====================================================
st.markdown("## 📤 Upload Excel")

uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel ที่กรอกข้อมูลแล้ว",
    type=["xlsx"]
)

if uploaded_file:
    with open("uploaded.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success(f"✅ อัปโหลดแล้ว: {uploaded_file.name}")

    # =================================================
    # GENERATE PAYLOAD
    # =================================================
    if st.button("🧩 Generate JSON Payload"):
        with st.spinner("⏳ กำลัง generate payload..."):
            files = generate_payload("uploaded.xlsx")

        if not files:
            st.warning("⚠️ ไม่พบข้อมูลที่สามารถ generate ได้")
        else:
            st.success(f"✅ Generate สำเร็จ {len(files)} ไฟล์")

            # zip download
            zip_path = "payloads.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for f in files:
                    zipf.write(f, arcname=os.path.basename(f))

            with open(zip_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Payloads (ZIP)",
                    f,
                    file_name="payloads.zip",
                    mime="application/zip"
                )

    st.divider()

# =====================================================
# SEND TO API
# =====================================================
if os.path.exists(PAYLOAD_DIR):
    st.markdown("## ☑️ Select Payload & Send to API")

    # -------- API AUTH --------
    access_token = st.text_input("🔑 Bearer Token", type="password")
    x_user = st.text_input("👤 x-user-name")

    if "sent_files" not in st.session_state:
        st.session_state.sent_files = set()

    payload_files = sorted(os.listdir(PAYLOAD_DIR))

    # -------- SELECT ALL --------
    col1, col2 = st.columns(2)
    if col1.button("✅ Select All"):
        for f in payload_files:
            if f not in st.session_state.sent_files:
                st.session_state[f] = True

    if col2.button("❌ Unselect All"):
        for f in payload_files:
            st.session_state[f] = False

    st.divider()

    selected_files = []

    # -------- FILE CHECKBOX LIST --------
    for filename in payload_files:
        disabled = filename in st.session_state.sent_files

        checked = st.checkbox(
            filename,
            key=filename,
            disabled=disabled
        )

        if checked and not disabled:
            selected_files.append(filename)

    # =================================================
    # SEND BUTTON
    # =================================================
    if st.button("🚀 Send to API", disabled=not selected_files):
        if not access_token or not x_user:
            st.error("❌ กรุณากรอก Bearer Token และ x-user-name")
        else:
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {access_token}",
                "x-user-name": x_user
            }

            results = []

            with st.spinner("⏳ กำลังยิง API ทีละไฟล์..."):
                for filename in selected_files:
                    file_path = os.path.join(PAYLOAD_DIR, filename)

                    with open(file_path, "r", encoding="utf-8") as f:
                        payload = json.load(f)

                    try:
                        resp = requests.post(
                            API_URL,
                            headers=headers,
                            json=payload,
                            timeout=30
                        )

                        result = {
                            "file": filename,
                            "status": resp.status_code,
                            "response": resp.text
                        }

                        if resp.ok:
                            st.session_state.sent_files.add(filename)

                    except Exception as e:
                        result = {
                            "file": filename,
                            "status": "ERROR",
                            "response": str(e)
                        }

                    results.append(result)

            # =================================================
            # RESULT SUMMARY
            # =================================================
            st.divider()
            st.markdown("## 📊 Result")

            failed = []

            for r in results:
                if r["status"] == "ERROR" or int(r["status"]) >= 400:
                    failed.append(r["file"])
                    st.error(f"❌ {r['file']} | {r['status']}")
                    st.code(r["response"])
                else:
                    st.success(f"✅ {r['file']} | {r['status']}")
                    st.code(r["response"])

            if failed:
                st.warning("⚠️ ไฟล์ที่ยิงไม่สำเร็จ:")
                st.write(failed)
            else:
                st.success("🎉 ยิง API สำเร็จทุกไฟล์")

st.divider()
st.caption("© Internal Tool | Excel → JSON → Disbursement API")
