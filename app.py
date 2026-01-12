import streamlit as st
import os
import zipfile
from generate_payload import generate_payload

# ================= PAGE CONFIG =================
st.set_page_config(
    page_title="API Payload Generator",
    layout="centered"
)

st.title("📄 API Payload Generator")
st.caption("Generate JSON payload from Excel (Row 22 header format)")

# ================= DOWNLOAD TEMPLATE =================
st.markdown("## 📥 Download Excel Template")

SAMPLE_FILE_PATH = "API_Transaction.xlsx"

if os.path.exists(SAMPLE_FILE_PATH):
    with open(SAMPLE_FILE_PATH, "rb") as f:
        st.download_button(
            label="⬇️ Download Excel Template",
            data=f,
            file_name="API_Transaction.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("⚠️ ไม่พบไฟล์ตัวอย่าง (API_Transaction.xlsx)")

st.info(
    "📌 กรุณาใช้ไฟล์นี้เป็น Template\n\n"
    "- Header ต้องอยู่ที่ Row 22\n"
    "- ข้อมูลเริ่ม Row 23\n"
    "- ห้ามเปลี่ยนลำดับ column"
)

st.divider()

# ================= UPLOAD FILE =================
st.markdown("## 📤 Upload Excel File")

uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel",
    type=["xlsx"]
)

if uploaded_file is not None:
    st.success(f"✅ อัปโหลดไฟล์: {uploaded_file.name}")

    # Save uploaded file temporarily
    temp_excel_path = "uploaded.xlsx"
    with open(temp_excel_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # ================= GENERATE =================
    if st.button("🚀 Generate Payload"):
        with st.spinner("⏳ กำลัง generate payload..."):
            try:
                output_dir = "payloads"

                generated_files = generate_payload(
                    excel_path=temp_excel_path,
                    sheet_name="API_Doc",
                    output_dir=output_dir,
                    debug=False
                )

                if not generated_files:
                    st.warning("⚠️ ไม่พบข้อมูลใน Excel")
                else:
                    st.success(f"✅ Generate สำเร็จ {len(generated_files)} ไฟล์")

                    # -------- ZIP FILE --------
                    zip_path = "payloads.zip"
                    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                        for file_path in generated_files:
                            zipf.write(
                                file_path,
                                arcname=os.path.basename(file_path)
                            )

                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label="⬇️ Download Payloads (ZIP)",
                            data=f,
                            file_name="payloads.zip",
                            mime="application/zip"
                        )

                    # -------- PREVIEW FIRST FILE --------
                    st.markdown("### 🔍 Preview Payload แรก")
                    with open(generated_files[0], "r", encoding="utf-8") as f:
                        st.json(f.read())

            except Exception as e:
                st.error("❌ เกิดข้อผิดพลาด")
                st.exception(e)

st.divider()

st.caption("© Internal Tool | Powered by Streamlit")
