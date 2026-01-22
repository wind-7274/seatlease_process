import streamlit as st
import msoffcrypto
import io
import zipfile

st.set_page_config(
    page_title="Excel Password Remover",
    page_icon="🔓"
)

st.title("🔓 Excel Password Remover")

# ----------------------------
# Clear / Reset function
# ----------------------------
def clear_all():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

# ----------------------------
# Inputs
# ----------------------------
password = st.text_input(
    "Enter Excel password",
    type="password",
    key="password_input"
)

uploaded_files = st.file_uploader(
    "Upload one or more .xlsx files",
    type=["xlsx"],
    accept_multiple_files=True,
    key="file_uploader"
)

# ----------------------------
# Buttons
# ----------------------------
col1, col2 = st.columns(2)

with col1:
    process_btn = st.button("Remove password")

with col2:
    st.button("Clear all", on_click=clear_all)

# ----------------------------
# Processing
# ----------------------------
if process_btn:
    if not uploaded_files:
        st.warning("Please upload at least one Excel file.")
    elif not password:
        st.warning("Please enter the password.")
    else:
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for uploaded_file in uploaded_files:
                try:
                    office_file = msoffcrypto.OfficeFile(uploaded_file)
                    office_file.load_key(password=password)

                    decrypted = io.BytesIO()
                    office_file.decrypt(decrypted)

                    zipf.writestr(uploaded_file.name, decrypted.getvalue())

                except Exception as e:
                    st.error(f"❌ Failed: {uploaded_file.name} ({e})")

        st.success("✅ Done!")

        st.download_button(
            "⬇ Download unprotected Excel files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="unprotected_excels.zip",
            mime="application/zip"
        )