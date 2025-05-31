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

    # Columns to merge on
    key_cols = ['Hierarchy Path', 'XML Tag']
    # Columns in source to compare
    source_cols = [
        col for col in source_df.columns
        if col not in key_cols + ['Level', 'Lvl']
        and not (isinstance(col, str) and col.startswith('Unnamed'))
    ]

    # Ensure all comparison columns are strings with no NaNs
    source_df[source_cols] = source_df[source_cols].fillna("").astype(str).applymap(str.strip)
    test_df[source_cols] = test_df[source_cols].fillna("").astype(str).applymap(str.strip)

    # Drop duplicates in source on keys
    source_clean = source_df[key_cols + source_cols].drop_duplicates(subset=key_cols)

    # Merge
    merged = pd.merge(
        test_df,
        source_clean,
        on=key_cols,
        how='left',
        suffixes=('', '_source')
    )

    # Build differences DataFrame
    differences = []
    for _, row in merged.iterrows():
        for col in source_cols:
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

    # Drop source columns from merged
    merged.drop(columns=[f"{col}_source" for col in source_cols if f"{col}_source" in merged.columns], inplace=True)

    # Clean merged output
    merged = merged.astype(str).replace("nan", "")
    merged = merged.replace({r'_x000D_': ' ', r'\r': ' ', r'\n': ' '}, regex=True)

    # Clean stripped source export
    stripped_source_export = source_df.copy()
    if 'Hierarchy Path' in stripped_source_export.columns:
        stripped_source_export.drop(columns=['Hierarchy Path'], inplace=True)
    stripped_source_export = stripped_source_export.replace("nan", "").replace({pd.NA: "", None: ""}).fillna("")
    stripped_source_export = stripped_source_export.replace({r'_x000D_': ' ', r'\r': ' ', r'\n': ' '}, regex=True)

    # Write to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in source_excel.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        stripped_source_export.to_excel(writer, sheet_name='Stripped Source', index=False)
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
