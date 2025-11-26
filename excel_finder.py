import streamlit as st
import os
import zipfile
from io import BytesIO

st.title("📦 Excel Finder + ZIP Extractor (Hybrid Local / Cloud)")

# ------------------------------
# User Inputs
# ------------------------------

st.subheader("1️⃣ Local Folder Search (optional)")

default_folder = r"C:\Users\Administrator\Downloads"
folder = st.text_input(
    "📁 Enter local folder path (leave blank if uploading files):", value=default_folder
)

st.subheader("2️⃣ Or Upload Files for Cloud / Remote Deployment")

uploaded_files = st.file_uploader(
    "Upload Excel or ZIP files",
    type=["xlsx", "xls", "xlsm", "zip"],
    accept_multiple_files=True,
)

patterns_input = st.text_area(
    "Enter filename keywords (partial or full):",
    placeholder="Example:\nsales\nPO123\nreport",
)

search_btn = st.button("🔍 Search Files")

excel_ext = (".xlsx", ".xls", ".xlsm")

# ------------------------------
# Search Logic
# ------------------------------
if search_btn:
    patterns = [p.strip().lower() for p in patterns_input.split("\n") if p.strip()]
    if not patterns:
        st.error("⚠️ Please enter at least one keyword.")
        st.stop()

    found_files = []

    # ------------------------------
    # 1. Local folder search
    # ------------------------------
    if folder and os.path.exists(folder):
        for root, dirs, files in os.walk(folder):
            for file in files:
                file_lower = file.lower()
                file_path = os.path.join(root, file)

                # Match Excel files
                if file_lower.endswith(excel_ext) and any(
                    p in file_lower for p in patterns
                ):
                    found_files.append(file_path)

                # Match Excel files inside ZIPs
                elif file_lower.endswith(".zip"):
                    try:
                        with zipfile.ZipFile(file_path, "r") as z:
                            for zip_item in z.namelist():
                                zip_item_lower = zip_item.lower()
                                if zip_item_lower.endswith(excel_ext) and any(
                                    p in zip_item_lower for p in patterns
                                ):
                                    data = z.read(zip_item)
                                    arcname = f"{os.path.splitext(file)[0]}_{os.path.basename(zip_item)}"
                                    found_files.append((arcname, data))
                    except zipfile.BadZipFile:
                        st.warning(f"⚠️ Cannot read ZIP file: {file}")

    # ------------------------------
    # 2. Uploaded files (cloud)
    # ------------------------------
    if uploaded_files:
        for file in uploaded_files:
            file_name = file.name
            file_bytes = file.read()

            if file_name.lower().endswith(excel_ext):
                if any(p in file_name.lower() for p in patterns):
                    found_files.append((file_name, file_bytes))
            elif file_name.lower().endswith(".zip"):
                try:
                    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
                        for zip_item in z.namelist():
                            zip_item_lower = zip_item.lower()
                            if zip_item_lower.endswith(excel_ext) and any(
                                p in zip_item_lower for p in patterns
                            ):
                                extracted_bytes = z.read(zip_item)
                                arcname = f"{file_name.split('.')[0]}_{zip_item.split('/')[-1]}"
                                found_files.append((arcname, extracted_bytes))
                except zipfile.BadZipFile:
                    st.warning(f"⚠️ Cannot read ZIP file: {file_name}")

    # ------------------------------
    # Show results
    # ------------------------------
    st.subheader("🔍 Search Results")
    if not found_files:
        st.error("❌ No matching Excel files found.")
        st.stop()

    for f in found_files:
        if isinstance(f, str):
            st.write(f"📄 {os.path.basename(f)}")
        else:
            st.write(f"🗜️ {f[0]}")

    # ------------------------------
    # Download all matches as ZIP
    # ------------------------------
    st.subheader("⬇️ Download All Matched Files (ZIP)")

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_out:
        for f in found_files:
            if isinstance(f, str):
                zip_out.write(f, os.path.basename(f))
            else:
                arcname, data = f
                zip_out.writestr(arcname, data)

    zip_buffer.seek(0)
    st.download_button(
        label="📦 Download ZIP",
        data=zip_buffer,
        file_name="excel_results.zip",
        mime="application/zip",
    )
