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

    # Columns to compare (exclude keys and structural columns)
    exclude_cols = ['Hierarchy Path', 'XML Tag', 'Level', 'Lvl']
    source_cols = [col for col in source_df.columns if col not in exclude_cols and not (isinstance(col, str) and col.startswith('Unnamed'))]

    key_cols = ['Hierarchy Path', 'XML Tag']

    # Normalize strings: replace nan/None with empty strings for consistent comparison
    for df in [source_df, test_df]:
        cols_to_process = [col for col in source_cols if col in df.columns]
        df[cols_to_process] = df[cols_to_process].fillna("").astype(str).applymap(str.strip)
        df[key_cols] = df[key_cols].fillna("").astype(str).applymap(str.strip)

    # Perform full outer merge on keys to capture all rows
    merged = pd.merge(
        source_df[key_cols + source_cols],
        test_df[key_cols + source_cols],
        on=key_cols,
        how='outer',
        suffixes=('_source', '_test'),
        indicator=True  # To track where row came from
    )

    # Function to detect changes per row
    def detect_change(row):
        if row['_merge'] == 'left_only':
            return 'Deleted in Test'
        elif row['_merge'] == 'right_only':
            return 'Added in Test'
        else:
            # Check if any column differs
            for col in source_cols:
                if row[f"{col}_source"] != row[f"{col}_test"]:
                    return 'Modified'
            return 'Unchanged'

    merged['Change Type'] = merged.apply(detect_change, axis=1)

    # Optional: filter out unchanged rows if you want
    # merged = merged[merged['Change Type'] != 'Unchanged']

    # Clean up _merge column as no longer needed
    merged.drop(columns=['_merge'], inplace=True)

    # Clean unwanted control characters globally
    for col in merged.columns:
        if merged[col].dtype == object:
            merged[col] = merged[col].str.replace(r'_x000D_|[\r\n]', ' ', regex=True)

    # Prepare stripped source for export
    stripped_source_export = source_df.copy()
    if 'Hierarchy Path' in stripped_source_export.columns:
        stripped_source_export.drop(columns=['Hierarchy Path'], inplace=True)
    stripped_source_export = stripped_source_export.replace({pd.NA: "", None: "", "nan": ""}).fillna("")
    stripped_source_export = stripped_source_export.replace({r'_x000D_': ' ', r'\r': ' ', r'\n': ' '}, regex=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write original source sheets
        for sheet_name, df in source_excel.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        stripped_source_export.to_excel(writer, sheet_name='Stripped Source', index=False)

        # Write merged output (test + matched source columns, without _test/_source suffixes)
        # We can write test_df as merged output here if you want:
        test_output_cols = [col for col in test_df.columns if col != 'Hierarchy Path']
        test_df.to_excel(writer, sheet_name='Merged Output', index=False)

        # Write the difference tracking sheet
        merged.to_excel(writer, sheet_name='Tracking Differences', index=False)

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
