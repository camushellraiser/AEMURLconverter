from openpyxl.styles import Alignment
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from urllib.parse import urlparse
import os

# Mapeo de idioma a ruta
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

def process_file(file_path):
    df = pd.read_excel(file_path, sheet_name=0, header=3)
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

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df_result = process_file(file_path)
            if df_result.empty:
                messagebox.showwarning("No Data", "No valid data found in the file.")
            else:
                                        output_path = os.path.splitext(file_path)[0] + "_converted.xlsx"
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                df_result.to_excel(writer, index=False)
                worksheet = writer.sheets["Sheet1"]
                worksheet.column_dimensions["A"].width = 60
                worksheet.column_dimensions["B"].width = 20
                worksheet.column_dimensions["C"].width = 80
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical="top")
                for cell in worksheet[1]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=2):
                    for cell in row:
                        cell.alignment = Alignment(horizontal="center", vertical="top")

            messagebox.showinfo("✅ Success", f"File saved to:\n{output_path}")
        except Exception as e:
            messagebox.showerror("❌ Error", str(e))

root = tk.Tk()
root.title("🌍 AEM - URL Converter")
root.geometry("400x200")
root.configure(bg="#f2f2f2")

button = tk.Button(root, text="📂 Browse Excel File", command=browse_file, height=2, width=30)
button.pack(pady=60)

root.mainloop()