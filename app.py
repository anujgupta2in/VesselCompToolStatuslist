import streamlit as st
import pandas as pd
from new_title_comparison import compare_titles
from comparison_utils import process_files
import io

# Set page config
st.set_page_config(
    page_title="Machinery Jobs Comparison",
    page_icon="🚢",
    layout="wide"
)

st.title("🚢 Machinery Jobs Comparison Tool")

st.markdown("""
This tool compares machinery jobs between two CSV files:
1. It analyzes job titles to identify differences for the same machinery
2. It compares job counts for each machinery item
3. It generates detailed Excel reports for both analyses

**Enhanced Features:**
- Color-coded display of job title differences:
  - **Green**: Common titles found in both files
  - **Orange**: Titles only found in the first file
  - **Blue**: Titles only found in the second file
  - **Purple**: Count columns
- Supports comparing files from the same vessel (adds file identifiers to columns)
- Improved visual organization with expandable sections
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Job List File")
    file1 = st.file_uploader("Upload Job List( System Management) CSV file", type=["csv"])

with col2:
    st.subheader("Second File")
    file2 = st.file_uploader("Upload Job Status CSV file", type=["csv"])

# Session state initialization
if 'title_diff_df' not in st.session_state:
    st.session_state.title_diff_df = None
if 'machinery_diff_list' not in st.session_state:
    st.session_state.machinery_diff_list = None
if 'title_excel_data' not in st.session_state:
    st.session_state.title_excel_data = None
if 'count_comparison_df' not in st.session_state:
    st.session_state.count_comparison_df = None
if 'count_excel_data' not in st.session_state:
    st.session_state.count_excel_data = None

if file1 and file2:
    try:
        file1_content = file1.getvalue()
        file2_content = file2.getvalue()

        with st.spinner("Processing files for both comparisons..."):
            title_diff_df, machinery_diff_list, title_excel_data = compare_titles(
                file1_content, file2_content, file1.name, file2.name
            )
            # Rename columns globally before saving to session and Excel
            title_diff_df = title_diff_df.rename(columns={
                title_diff_df.columns[3]: 'Titles only in Job List File',
                title_diff_df.columns[4]: 'Titles only in Job Status File'
            })

            count_comparison_df, count_excel_data = process_files(
                file1_content, file2_content, file1.name, file2.name
            )

            st.session_state.title_diff_df = title_diff_df
            st.session_state.machinery_diff_list = machinery_diff_list
            st.session_state.title_excel_data = title_excel_data
            st.session_state.count_comparison_df = count_comparison_df
            st.session_state.count_excel_data = count_excel_data

            st.success("Files processed successfully! View results in the tabs below.")
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        st.exception(e)

# Tabs
tab1, tab2 = st.tabs(["Job Title Comparison", "Machinery Count Comparison"])

with tab1:
    st.header("Job Title Comparison Results")
    if st.session_state.title_diff_df is not None and st.session_state.machinery_diff_list is not None:
        title_diff_df = st.session_state.title_diff_df
        machinery_diff_list = st.session_state.machinery_diff_list
        title_excel_data = st.session_state.title_excel_data

        st.subheader("📊 Comparison Summary")
        total_machinery = len(title_diff_df)
        diff_count = len(machinery_diff_list)
        same_count = total_machinery - diff_count

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Machinery Items", total_machinery)
        col2.metric("Items with Different Titles", diff_count)
        col3.metric("Items with Same Titles", same_count)

        if diff_count > 0:
            st.subheader("📋 Machinery with Different Job Titles")
            st.write(f"There are **{diff_count}** machinery items with different job titles:")
            st.text_area("Machinery List:", "\n".join([f"• {m}" for m in machinery_diff_list]), height=150)

            st.subheader("🔄 Detailed Title Comparison")
            diff_only_df = title_diff_df[title_diff_df['Has Differences'] == 'Yes'].copy()
            # Rename columns for user-friendly labeling
            diff_only_df = diff_only_df.rename(columns={
                diff_only_df.columns[3]: 'Titles only in Job List File',
                diff_only_df.columns[4]: 'Titles only in Job Status File'
            })

            title_cols = diff_only_df.columns[2:5]  # Common Titles, Titles only in File 1, Titles only in File 2
            for col in title_cols:
                def format_with_count(row):
                    if row[col] == '-' or pd.isna(row[col]):
                        return row[col]
                    count = len([x for x in row[col].split(', ') if x.strip()])
                    return f"{row[col]}\n(count: {count})" if count > 0 else row[col]
                diff_only_df[col] = diff_only_df.apply(format_with_count, axis=1)

            def highlight_title_counts(row):
                styles = [''] * len(row)
                for idx, col in enumerate(row.index):
                    if 'Titles only in' in col and row[col] != '-':
                        styles[idx] = 'background-color: #FFF3E0'  # orange
                    elif 'Common Titles' in col and row[col] != '-':
                        styles[idx] = 'background-color: #E8F5E9'  # green
                    elif 'Count' in col and row[col] > 0:
                        styles[idx] = 'background-color: #E3E1F7'  # purple tint
                return styles

            styled_df = diff_only_df.style.apply(highlight_title_counts, axis=1)
            st.dataframe(styled_df, use_container_width=True)

            if isinstance(title_excel_data, bytes):
                st.subheader("📅 Download Report")
                st.download_button(
                    label="Download Job Title Comparison Report",
                    data=title_excel_data,
                    file_name="Job_Title_Comparison.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.success("No job title differences found for any machinery!")
    else:
        st.info("Please upload both CSV files to generate the title comparison report.")

with tab2:
    st.header("Machinery Count Comparison Results")
    if st.session_state.count_comparison_df is not None:
        comparison_df = st.session_state.count_comparison_df
        excel_data = st.session_state.count_excel_data

        def highlight_differences(row):
            styles = [''] * len(row)
            if row['Machinery'] != 'TOTAL':
                col1_name = comparison_df.columns[1]
                col2_name = comparison_df.columns[2]
                if row[col1_name] == 0 or row[col2_name] == 0:
                    styles[0] = 'background-color: #FFC7CE; font-weight: bold'
                    styles[3] = 'background-color: #FFC7CE; color: #9C0006'
                if row[col1_name] != row[col2_name]:
                    styles[1] = 'background-color: #FFEB9C'
                    styles[2] = 'background-color: #FFEB9C'
                    if row[col1_name] > row[col2_name]:
                        styles[3] = 'background-color: #C6EFCE; color: #006100'
                    else:
                        styles[3] = 'background-color: #FFC7CE; color: #9C0006'
            else:
                return ['font-weight: bold'] * len(row)
            return styles

        styled_df = comparison_df.style.apply(highlight_differences, axis=1)
        st.dataframe(styled_df, use_container_width=True)

        st.download_button(
            label="Download Excel Report",
            data=excel_data,
            file_name="Machinery_Count_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.info("""
        **Explanation:**
        - **Red highlighting**: Machinery that only exists in one file
        - **Yellow highlighting**: Different job counts between files
        - **Green (positive difference)**: More jobs in first file
        - **Red (negative difference)**: More jobs in second file

        Note: The color coding is applied to both the online view and the Excel report.
        """)
    else:
        st.info("Please upload both CSV files to generate the machinery count comparison report.")
