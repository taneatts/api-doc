import streamlit as st
import os
import zipfile
from generate_payload import generate_payload

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="API Payload Generator",
    page_icon="📄",
    layout="centered"
)

# =========================================================
# HEADER
# =========================================================
st.markdown(
    """
    <h2 style="text-align:center;">📄 API Payload Generator</h2>
    <p style="text-align:center; color:gray;">
    Generate JSON payload from Excel Template (Agent / Broker / Company)
    </p>
    """,
    unsafe_allow_html=True
)

st.divider()

# =========================================================
# STEP 1 : DOWNLOAD TEMPLATE
# =========================================================
with st.container():
    st.markdown("### 🧩 Step 1: Download Excel Template")

    col1, col2 = st.columns([1, 2])

    with col1:
        st.markdown("**📥 Template File**")

    with col2:
        TEMPLATE_FILE = "API_Transaction.xlsx"

        if os.path.exists(TEMPLATE_FILE):
            with open(TEMPLATE_FILE, "rb") as f:
                st.download_button(
                    label="⬇️ Download Excel Template (Current Version)",
                    data=f,
                    file_name="API_Transaction.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("❌ ไม่พบไฟล์ Template (API_Transaction.xlsx)")

    with st.expander("📌 วิธีใช้งาน Excel Template"):
        st.markdown(
            """
            #### 1️⃣ โครงสร้างชีท
            - `API_Doc_Agent_Broker`
            - `API_Doc_Company`

            #### 2️⃣ Header / Data
            - Header อยู่ที่ **Row 22**
            - ข้อมูลเริ่มที่ **Row 23**
            - Payload เริ่มที่ **Column E**
            - ❌ ห้ามเปลี่ยนลำดับ Column

            #### 3️⃣ การตั้งชื่อไฟล์ JSON (Column A–D)
            | Column | ความหมาย |
            |------|---------|
            | A | ประเภทงาน |
            | B | ประเภทผู้รับเงิน |
            | C | วิธีจ่ายเงิน |
            | D | Running No |

            **ตัวอย่างชื่อไฟล์**
            ```
            GCM ค่านายหน้า_Agent_Bank transfer_DT0001.json
            ```

            #### 4️⃣ เงื่อนไขสำคัญ
            - `tax` ว่าง → `null`
            - `committees`
              - Agent/Broker → แสดงค่าตามข้อมูล หรือ `null`
              - Company → อ่านจาก column ที่กำหนด
            """
        )

st.divider()

# =========================================================
# STEP 2 : UPLOAD FILE
# =========================================================
with st.container():
    st.markdown("### 📤 Step 2: Upload Excel File")

    uploaded_file = st.file_uploader(
        "เลือกไฟล์ Excel ที่กรอกข้อมูลแล้ว",
        type=["xlsx"]
    )

    if uploaded_file:
        st.success(f"✅ อัปโหลดไฟล์: {uploaded_file.name}")

        temp_excel_path = "uploaded.xlsx"
        with open(temp_excel_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

st.divider()

# =========================================================
# STEP 3 : GENERATE PAYLOAD
# =========================================================
with st.container():
    st.markdown("### 🚀 Step 3: Generate Payload")

    if uploaded_file:
        if st.button("Generate JSON Payload", use_container_width=True):
            with st.spinner("⏳ กำลัง generate payload จากทั้ง 2 ชีท..."):
                try:
                    output_dir = "payloads"

                    generated_files = generate_payload(
                        excel_path=temp_excel_path,
                        output_dir=output_dir,
                        debug=False
                    )

                    if not generated_files:
                        st.warning("⚠️ ไม่พบข้อมูลที่สามารถ generate ได้")
                    else:
                        st.success(f"✅ Generate สำเร็จทั้งหมด {len(generated_files)} ไฟล์")

                        # ZIP FILE
                        zip_path = "payloads.zip"
                        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                            for file_path in generated_files:
                                zipf.write(
                                    file_path,
                                    arcname=os.path.basename(file_path)
                                )

                        st.download_button(
                            label="⬇️ Download Payloads (ZIP)",
                            data=open(zip_path, "rb"),
                            file_name="payloads.zip",
                            mime="application/zip",
                            use_container_width=True
                        )

                        # PREVIEW
                        st.markdown("#### 🔍 Preview ตัวอย่าง Payload")
                        with open(generated_files[0], "r", encoding="utf-8") as f:
                            st.json(f.read())

                except Exception as e:
                    st.error("❌ เกิดข้อผิดพลาดระหว่าง generate payload")
                    st.exception(e)
    else:
        st.info("ℹ️ กรุณาอัปโหลดไฟล์ Excel ก่อน")

st.divider()

st.caption("© Internal Tool | Excel → JSON Payload Generator")
