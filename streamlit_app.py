import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.title("📞 Excel Splitter, Auto-Formatter & Phone Validator")

st.write("""
Upload an Excel file with two columns:
- **ID** — identifier (kept exactly as in your raw file)  
- **Numbers** — list of phone numbers separated by a character (like `,` or `;`)  

The app will:
1. Split numbers by your chosen separator  
2. Auto-format them (mobile or landline), including `*0`, `*09`, and `*63` numbers  
3. Validate only correct phone numbers:
   - ✅ Mobile: `0XXXXXXXXXX` (11 digits, starts with 0)
   - ✅ Landline: `0XXXXXXXXX` (10 digits, starts with 0)
   - ❌ Invalid if starting with `00` or `3`
4. Show one row per ID with valid numbers only under headers **TU1, TU2, TU3, ...**
5. Allow download of **valid results** and **invalid numbers**
""")

uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx", "xls"])
separator = st.text_input("Enter the separator (e.g., ',' or ';' or '|')", value=",")

# Helper functions
def to_str(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def clean_value(val):
    """Remove all non-digit characters except '*'."""
    val = to_str(val)
    return re.sub(r"[^\d*]", "", val)

def auto_format_number(val):
    """Auto-format PH phone numbers."""
    val = clean_value(val)
    if val == "":
        return ""

    if val.startswith("*"):
        val = val[1:]

    if val.startswith("63") and len(val) >= 12:
        return "0" + val[-10:]

    if re.fullmatch(r"9\d{9}", val):
        return "0" + val

    if re.fullmatch(r"0\d{10}", val):
        return val

    if re.fullmatch(r"0\d{9}", val):
        return val

    if re.fullmatch(r"\d{10}", val):
        return val

    if re.fullmatch(r"\d{7,9}", val):
        return "0" + val

    if re.fullmatch(r"\d{11}", val):
        return "0" + val[1:]

    return val

def is_valid_phone(val):
    val = to_str(val)

    if val.startswith("00") or val.startswith("3"):
        return False

    pattern_mobile = re.compile(r"^0\d{10}$")
    pattern_landline = re.compile(r"^0\d{9}$")

    return bool(pattern_mobile.match(val) or pattern_landline.match(val))

# App logic
if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, dtype=str)
        st.write("### 📄 Preview of uploaded file:")
        st.dataframe(df_raw.head())

        if df_raw.shape[1] < 2:
            st.error("❌ The Excel file must have at least two columns: ID and Numbers.")
        else:
            id_col = df_raw.columns[0]
            num_col = df_raw.columns[1]

            all_valid = []
            all_invalid = []

            for _, row in df_raw.iterrows():
                raw_id = row[id_col]
                raw_numbers = to_str(row[num_col])

                numbers = [n.strip() for n in raw_numbers.split(separator) if n.strip() != ""]

                valid_nums = []
                invalid_nums = []

                for num in numbers:
                    formatted = auto_format_number(num)
                    if is_valid_phone(formatted):
                        valid_nums.append(formatted)
                    else:
                        invalid_nums.append(num)

                valid_row = {id_col: raw_id}
                for i, num in enumerate(valid_nums, start=1):
                    valid_row[f"TU{i}"] = num
                all_valid.append(valid_row)

                for bad in invalid_nums:
                    all_invalid.append({id_col: raw_id, "Invalid Value": bad})

            valid_df = pd.DataFrame(all_valid).fillna("")
            invalid_df = pd.DataFrame(all_invalid).fillna("")

            st.success("✅ Processing complete! One row per ID with valid numbers only.")
            st.write(f"**✅ Valid IDs: {len(valid_df)} | ❌ Invalid entries: {len(invalid_df)}**")

            st.write("### ✅ Valid Numbers (One Row per ID)")
            st.dataframe(valid_df)

            st.write("### ❌ Invalid Numbers")
            st.dataframe(invalid_df)

            valid_output = BytesIO()
            valid_df.to_excel(valid_output, index=False)
            st.download_button(
                label="💾 Download VALID Numbers (1 row per ID)",
                data=valid_output.getvalue(),
                file_name="valid_numbers_per_id.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if not invalid_df.empty:
                invalid_output = BytesIO()
                invalid_df.to_excel(invalid_output, index=False)
                st.download_button(
                    label="💾 Download INVALID Numbers",
                    data=invalid_output.getvalue(),
                    file_name="invalid_numbers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
