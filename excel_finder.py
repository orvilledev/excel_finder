import streamlit as st
import os
import zipfile
from io import BytesIO

st.title("📦 Excel Finder + ZIP Extractor (Local Folder)")

# Default folder path (editable)
default_folder = r"C:\Users\Administrator\Downloads"
folder = st.text_input("📁 Enter folder path:", value=default_folder)

# Filename keywords
patterns_input = st.text_area(
    "Enter keywords to search in filenames (partial or full):",
    placeholder="Example:\nsales\nPO123\nreport",
)

search_btn = st.button("🔍 Search Files")

excel_ext = (".xlsx", ".xls", ".xlsm")

if search_btn:
    if not os.path.exists(folder):
        st.error("❌ Folder does not exist.")
        st.stop()

    patterns = [p.strip().lower() for p in patterns_input.split("\n") if p.strip()]
    if not patterns:
        st.error("⚠️ Please enter at least one keyword.")
        st.stop()

    found_files = []

    # -----------------------------
    # Search folder recursively
    # -----------------------------
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

    st.subheader("🔍 Search Results")

    if not found_files:
        st.error("❌ No matching Excel files found.")
        st.stop()

    for f in found_files:
        if isinstance(f, str):
            st.write(f"📄 {os.path.basename(f)}")
        else:
            st.write(f"🗜️ {f[0]}")

    # -----------------------------
    # Create ZIP for download
    # -----------------------------
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
