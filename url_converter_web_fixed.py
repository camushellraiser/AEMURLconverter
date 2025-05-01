
import streamlit as st
import pandas as pd
from urllib.parse import urlparse
from io import BytesIO
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl import load_workbook

st.set_page_config(page_title="AEM URL Converter", page_icon="🌍", layout="wide")

st.markdown("<h1 style='text-align: center; color: #2E86C1;'>🌍 AEM URL Converter</h1>", unsafe_allow_html=True)
st.markdown("Upload a Web Translation Excel file to convert AEM URLs based on target languages.")

LANGUAGE_MAP = {
    "de-DE": "/content/lifetech/europe/en-de",
    "es-ES": "/content/lifetech/europe/en-es",
    "fr-FR": "/content/lifetech/europe/en-fr",
    "ja-JP": "/content/lifetech/japan/en-jp",
    "ko-KR": "/content/lifetech/ipac/en-kr",
    "zh-CN": "/content/lifetech/greater-china/en-cn",
    "zh-TW": "/content/lifetech/ipac/en-tw",
    "pt-BR": "/content/lifetech/latin-america/en-br",
    "es-LATAM": "/content/lifetech/latin-america/en-mx"
}

def clean_url(url):
    if not isinstance(url, str):
        return None
    parsed = urlparse(url)
    path = parsed.path
    if "/home/" not in path:
        return None
    cleaned = path.split("/home/", 1)[1]
    if ".html" in cleaned:
        cleaned = cleaned.replace(".html", "")
    return "/home/" + cleaned

def detect_first_url(row):
    for cell in row:
        if isinstance(cell, str) and "/home/" in cell:
            return cell
    return None

def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0, header=3)
    results = []

    language_columns = {col: code for code, path in LANGUAGE_MAP.items() for col in df.columns if code in str(col)}

    for _, row in df.iterrows():
        original_url = detect_first_url(row)
        cleaned_path = clean_url(original_url)
        if not cleaned_path:
            continue

        for col_name, lang_code in language_columns.items():
            cell_value = row.get(col_name, "")
            if pd.notna(cell_value) and str(cell_value).strip().lower() not in ["", "no"]:
                localized_base = LANGUAGE_MAP.get(lang_code)
                if localized_base:
                    results.append({
                        "Original URL": original_url,
                        "Language": lang_code,
                        "Localized Path": localized_base + cleaned_path
                    })

    result_df = pd.DataFrame(results)
    return result_df.sort_values(by=["Language"]) if not result_df.empty else pd.DataFrame()

uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df_result = process_file(uploaded_file)
        if df_result.empty:
            st.warning("⚠️ No valid data found in the uploaded file.")
        else:
            st.success("✅ URLs converted successfully!")
            st.markdown("### 🔗 Localized URLs (select and copy supported)")

            # Apply styling for web view
            styled_df = df_result.style.set_properties(subset=["Language"], **{"text-align": "center"})
            st.dataframe(styled_df, use_container_width=True)

            # Prepare Excel file with formatting
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_result.to_excel(writer, index=False)
                worksheet = writer.sheets["Sheet1"]
                worksheet.freeze_panes = "A2"

                worksheet.column_dimensions["A"].width = 60  # Original URL
                worksheet.column_dimensions["B"].width = 15  # Language
                worksheet.column_dimensions["C"].width = 65  # Localized Path

                header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                header_font = Font(bold=True)

                for row_idx, row in enumerate(worksheet.iter_rows(), start=1):
                    for col_idx, cell in enumerate(row, start=1):
                        cell.alignment = Alignment(wrap_text=True, vertical="top")
                        if row_idx == 1:
                            cell.fill = header_fill
                            cell.font = header_font
                            if col_idx == 2:
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                        elif col_idx == 2:
                            cell.alignment = Alignment(horizontal="center", vertical="top")

            st.download_button(
                label="📥 Download Converted Excel",
                data=output.getvalue(),
                file_name="converted_urls.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
