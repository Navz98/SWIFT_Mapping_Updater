import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Mapping Sheet Updater", layout="wide")
st.title("SWIFT Mapping Sheet Updater")

def strip_all_string_columns(df):
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
    return df

def build_path_column(df):
    path_stack = {}
    full_paths = []

    for _, row in df.iterrows():
        level = row.get('Lvl')
        name = str(row.get('Name', '')).strip()
        tag = str(row.get('XML Tag', '')).strip()

        if pd.isna(level):
            full_paths.append(None)
            continue

        level = int(level)
        component = f"{tag}__{name}"
        path_stack[level] = component

        for lvl in list(path_stack.keys()):
            if lvl > level:
                del path_stack[lvl]

        sorted_levels = sorted(path_stack.keys())
        path = " > ".join([path_stack[lvl] for lvl in sorted_levels])
        full_paths.append(path)

    df['Hierarchy Path'] = full_paths
    return df

def process_excel(source_file, test_file):
    source_excel = pd.read_excel(source_file, sheet_name=None)
    test_excel = pd.read_excel(test_file, sheet_name=None)

    source_df = pd.concat(source_excel.values(), ignore_index=True)
    test_df = pd.concat(test_excel.values(), ignore_index=True)

    source_df = strip_all_string_columns(source_df)
    test_df = strip_all_string_columns(test_df)

    source_df = build_path_column(source_df)
    test_df = build_path_column(test_df)

    key_cols = ['Hierarchy Path', 'XML Tag']
    excluded_cols = key_cols + ['Level', 'Lvl']

    source_output_columns = [col for col in source_df.columns if col not in excluded_cols and not col.startswith('Unnamed')]
    test_output_columns = [col for col in test_df.columns if col not in excluded_cols and not col.startswith('Unnamed')]

    merge_columns = key_cols + source_output_columns

    source_clean = source_df[merge_columns].drop_duplicates(subset=key_cols)
    test_clean = test_df.copy()

    source_clean = source_clean.fillna("").astype(str).replace("nan", "")
    test_clean = test_clean.fillna("").astype(str).replace("nan", "")

    merged = pd.merge(test_clean, source_clean, on=key_cols, how='left', suffixes=('', '_source'))

    differences = []

    # 1. Cell-level differences (using test sheet columns only)
    for _, row in merged.iterrows():
        for col in test_output_columns:
            test_val = str(row.get(col, "")).strip()
            source_val = str(row.get(f"{col}_source", "")).strip()
            if test_val != source_val:
                differences.append({
                    "Hierarchy Path": row.get("Hierarchy Path", ""),
                    "XML Tag": row.get("XML Tag", ""),
                    "Column": col,
                    "Test Value": test_val,
                    "Source Value": source_val,
                    "Type": "Changed"
                })

    # 2. New rows in test
    source_keys = set(zip(source_clean['Hierarchy Path'], source_clean['XML Tag']))
    for _, row in test_clean.iterrows():
        key = (row['Hierarchy Path'], row['XML Tag'])
        if key not in source_keys:
            for col in test_output_columns:
                differences.append({
                    "Hierarchy Path": row['Hierarchy Path'],
                    "XML Tag": row['XML Tag'],
                    "Column": col,
                    "Test Value": row.get(col, ""),
                    "Source Value": "",
                    "Type": "New in Test"
                })

    # 3. Missing rows in test
    test_keys = set(zip(test_clean['Hierarchy Path'], test_clean['XML Tag']))
    for _, row in source_clean.iterrows():
        key = (row['Hierarchy Path'], row['XML Tag'])
        if key not in test_keys:
            for col in test_output_columns:
                differences.append({
                    "Hierarchy Path": row['Hierarchy Path'],
                    "XML Tag": row['XML Tag'],
                    "Column": col,
                    "Test Value": "",
                    "Source Value": row.get(col, ""),
                    "Type": "Missing in Test"
                })

    differences_df = pd.DataFrame(differences)

    # Prepare merged output
    merged.drop(columns=[f"{col}_source" for col in source_output_columns if f"{col}_source" in merged.columns], inplace=True)
    if 'Hierarchy Path' in merged.columns:
        merged.drop(columns=['Hierarchy Path'], inplace=True)
    merged = merged.astype(str).replace("nan", "")
    merged = merged.replace({r'_x000D_': ' ', r'\r': ' ', r'\n': ' '}, regex=True)

    # Clean source for export
    stripped_source_export = source_df.copy()
    if 'Hierarchy Path' in stripped_source_export.columns:
        stripped_source_export.drop(columns=['Hierarchy Path'], inplace=True)
    stripped_source_export = stripped_source_export.replace("nan", "").replace({pd.NA: "", None: ""}).fillna("")
    stripped_source_export = stripped_source_export.replace({r'_x000D_': ' ', r'\r': ' ', r'\n': ' '}, regex=True)

    # Write to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        stripped_source_export.to_excel(writer, sheet_name='Source', index=False)
        merged.to_excel(writer, sheet_name='New Mapping', index=False)
        if not differences_df.empty:
            differences_df.to_excel(writer, sheet_name='Differences', index=False)

    output.seek(0)
    return output

# Streamlit UI
source_file = st.file_uploader("⬆️ Upload Latest Mapping Excel File", type=[".xlsx"])
test_file = st.file_uploader("⬆️ Upload SWIFT Excel File", type=[".xlsx"])

if source_file and test_file:
    if st.button("Do the trick ✨"):
        with st.spinner("🥁 Drum Rolls..."):
            result = process_excel(source_file, test_file)
            st.success("Ta Da! Click the below button to download.")
            st.download_button("📥 Download Updated Mapping Sheet", result, file_name="Updated_mapping_sheet.xlsx")

# Footer
st.markdown("""
    <style>
    .footer {
        position: fixed;
        bottom: 10px;
        width: 100%;
        text-align: center;
        color: grey;
        font-size: 0.9rem;
        font-family: 'Courier New', monospace;
    }
    </style>
    <div class="footer">
         🧘‍♂️ Designed By Naveen 🧘‍♂️
    </div>
""", unsafe_allow_html=True)
