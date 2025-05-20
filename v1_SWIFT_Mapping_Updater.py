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

    source_output_columns = [
        col for col in source_df.columns
        if col not in ['Hierarchy Path', 'XML Tag', 'Level', 'Lvl'] and not (isinstance(col, str) and col.startswith('Unnamed'))
    ]

    merge_columns = ['Hierarchy Path', 'XML Tag'] + source_output_columns
    source_clean = source_df[merge_columns].drop_duplicates(subset=['Hierarchy Path', 'XML Tag'])

    merged = pd.merge(
        test_df,
        source_clean,
        on=['Hierarchy Path', 'XML Tag'],
        how='left',
        suffixes=('', '_source')
    )

    for i, row in merged.iterrows():
        xml_tag = row.get('XML Tag')
        hierarchy = row.get('Hierarchy Path')

        if pd.notna(xml_tag) and str(xml_tag).strip() != "":
            continue
        if pd.isna(hierarchy) or str(hierarchy).strip() == "":
            continue

        fallback_rows = source_df[
            (source_df['Hierarchy Path'] == hierarchy) &
            (pd.isna(source_df['XML Tag'])) &
            (source_df['Lvl'] == row['Lvl']) &
            (source_df['Name'] == row['Name'])
        ]

        if not fallback_rows.empty:
            fallback_row = fallback_rows.iloc[0]
            for col in source_output_columns:
                fallback_val = fallback_row.get(col)
                current_val = merged.at[i, col]
                if (pd.isna(current_val) or current_val == "") and pd.notna(fallback_val):
                    merged.at[i, col] = fallback_val

    final_columns_order = [col for col in source_df.columns if col != 'Hierarchy Path']
    final_columns_order = [col for col in final_columns_order if col in merged.columns]
    merged = merged[final_columns_order]

    merged = merged.astype(str).replace("nan", "")
    merged = merged.replace({r'_x000D_': ' ', r'\r': ' ', r'\n': ' '}, regex=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in source_excel.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        merged.to_excel(writer, sheet_name='Merged Output', index=False)

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

# Footer with trademark text centered at the bottom
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
