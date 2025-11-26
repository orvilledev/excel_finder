import streamlit as st
import os
import zipfile
from io import BytesIO

st.title("📦 Excel Finder (Folder + ZIP Scan)")

# Default folder path
default_path = r"C:\Users\Administrator\Downloads"
folder = st.text_input("📁 Folder to search:", value=default_path)

# User search patterns
patterns_input = st.text_area(
    "🔎 Enter filename keywords (partial or full):",
    placeholder="Example:\nsales\nPO123\nreport",
)

search_btn = st.button("Search Files")

# Allowed Excel extensions
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
    # SEARCH NORMAL FOLDERS
    # -----------------------------
    for root, dirs, files in os.walk(folder):
        for file in files:
            file_lower = file.lower()

            # Match Excel file normally
            if file_lower.endswith(excel_ext) and any(
                p in file_lower for p in patterns
            ):
                found_files.append(os.path.join(root, file))

            # Check inside .zip files
            if file_lower.endswith(".zip"):
                zip_path = os.path.join(root, file)
                try:
                    with zipfile.ZipFile(zip_path, "r") as zipf:
                        for zip_item in zipf.namelist():
                            zip_item_lower = zip_item.lower()
                            if zip_item_lower.endswith(excel_ext) and any(
                                p in zip_item_lower for p in patterns
                            ):
                                # Extract to memory
                                extracted = zipf.read(zip_item)
                                found_files.append((file, zip_item, extracted))
                except:
                    pass

    st.subheader("🔍 Results")

    if not found_files:
        st.error("❌ No matching Excel files found.")
        st.stop()

    # Show normal and ZIP results
    for item in found_files:
        if isinstance(item, str):
            st.write("📄 " + os.path.basename(item))
        else:
            zip_name, inner_file, _ = item
            st.write(f"🗜️ {zip_name} → {inner_file}")

    # -----------------------------
    # CREATE ZIP FOR DOWNLOAD
    # -----------------------------
    st.subheader("⬇️ Download All Matched Files (ZIP)")

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_out:
        for item in found_files:
            if isinstance(item, str):
                # Real file from folder
                zip_out.write(item, os.path.basename(item))
            else:
                # File extracted from ZIP archive
                zip_name, inner_file, data = item
                arcname = (
                    f"{os.path.splitext(zip_name)[0]}_{os.path.basename(inner_file)}"
                )
                zip_out.writestr(arcname, data)

    zip_buffer.seek(0)

    st.download_button(
        label="📦 Download ZIP",
        data=zip_buffer,
        file_name="excel_results.zip",
        mime="application/zip",
    )
