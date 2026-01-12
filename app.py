import streamlit as st
import os
import zipfile
from generate_payload import generate_payload

st.set_page_config(page_title="Payload Generator", layout="centered")

st.title("📄 API Payload Generator")
st.write("อัปโหลด Excel แล้วระบบจะ generate JSON payload ให้")

uploaded_file = st.file_uploader(
    "Upload API_Transaction.xlsx",
    type=["xlsx"]
)

if uploaded_file:
    with open("uploaded.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button("🚀 Generate Payload"):
        with st.spinner("กำลัง generate payload..."):
            files = generate_payload("uploaded.xlsx")

        zip_name = "payloads.zip"
        with zipfile.ZipFile(zip_name, "w") as zipf:
            for file in files:
                zipf.write(file, arcname=os.path.basename(file))

        st.success(f"✅ Generate สำเร็จ {len(files)} ไฟล์")

        with open(zip_name, "rb") as f:
            st.download_button(
                label="⬇️ Download Payloads (ZIP)",
                data=f,
                file_name="payloads.zip",
                mime="application/zip"
            )
