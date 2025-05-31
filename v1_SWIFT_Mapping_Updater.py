import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Mapping Sheet Updater", layout="wide")
st.title("SWIFT Mapping Sheet Updater")

# Helper Functions
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

    # Determine output columns excluding structural
    source_output_columns = [
        col for col in source_df.columns
        if col not in ['Hierarchy Path', 'XML Tag', 'Level', 'Lvl']
        and not (isinstance(col, str) and col.startswith('Unnamed'))
    ]

    key_cols = ['Hierarchy Path', 'XML Tag']
    source_clean = source_df[key_cols + source_output_columns].drop_duplicates(subset=key_cols)

    # Merge test data with source data
    merged = pd.merge(
        test_df,
        source_clean,
        on=key_cols,
        how='left',
        suffixes=('', '_source')
    )

    # Build differences
    differences = []
    merged.fillna("", inplace=True)
    for _, row in merged.iterrows():
        for col in source_output_columns:
            test_val = str(row.get(col, "")).strip()
            source_val = str(row.get(f"{col}_source", "")).strip()
            if test_val != source_val:
                differences.append({
                    "Hierarchy Path": row.get("Hierarchy Path", ""),
                    "XML Tag": row.get("XML Tag", ""),
                    "Column": col,
                    "Test Value": test_val,
                    "Source Value": source_val
                })

    differences_df = pd.DataFrame(differences)

    # Clean up merged dataframe
    merged.drop(columns=[f"{col}_source" for col in source_output_columns if f"{col}_source" in merged.columns], inplace=True)

    # Preserve column order of test_df
    test_columns_order = [col for col in test_df.columns if col in merged.columns]
    additional_cols = [col for col in merged.columns if col not in test_columns_order and col not in ['Hierarchy Path']]
    merged = merged[test_columns_order + additional_cols]

    # Remove 'Hierarchy Path' from merged output (was not there originally)
    if 'Hierarchy Path' in merged.columns:
        merged.drop(columns=['Hierarchy Path'], inplace=True)

    merged = merged.astype(str).replace("nan", "")
    merged = merged.replace({r'_x000D_': ' ', r'\r': ' ', r'\n': ' '}, regex=True)

    # Stripped and renamed source export
    stripped_source_export = source_df.copy()
    if 'Hierarchy Path' in stripped_source_export.columns:
        stripped_source_export.drop(columns=['Hierarchy Path'], inplace=True)
    stripped_source_export = stripped_source_export.replace("nan", "").replace({pd.NA: "", None: ""}).fillna("")
    stripped_source_export = stripped_source_export.replace({r'_x000D_': ' ', r'\r': ' ', r'\n': ' '}, regex=True)

    # Write to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        stripped_source_export.to_excel(writer, sheet_name='Source', index=False)
        merged.to_excel(writer, sheet_name='Merged Output', index=False)
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
