import streamlit as st
import pandas as pd
from new_title_comparison import compare_titles
from comparison_utils import process_files, build_frequency_comparison
import io

st.set_page_config(
    page_title="Job List and Job Status Page Comparison",
    page_icon="🚢",
    layout="wide"
)

st.title("🚢 Job List and Job Status Page Tool")

st.markdown("""
This tool compares machinery jobs between two CSV files:
1. It analyzes job titles to identify differences for the same machinery
2. It compares job counts for each machinery item
3. It generates detailed Excel reports for both analyses

**Color coding:**
- **Green**: Common titles found in both files
- **Orange**: Titles only found in the first file
- **Blue**: Titles only found in the second file
- **Purple**: Count columns
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Job List File")
    file1 = st.file_uploader("Upload Job List (System Management) CSV file", type=["csv"])

with col2:
    st.subheader("Second File")
    file2 = st.file_uploader("Upload Job Status CSV file", type=["csv"])

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
if 'job_detail' not in st.session_state:
    st.session_state.job_detail = None

if file1 and file2:
    try:
        file1_content = file1.getvalue()
        file2_content = file2.getvalue()

        with st.spinner("Processing files for both comparisons..."):
            title_diff_df, machinery_diff_list, title_excel_data = compare_titles(
                file1_content, file2_content, file1.name, file2.name
            )

            count_comparison_df, count_excel_data, job_detail = process_files(
                file1_content, file2_content, file1.name, file2.name
            )

            # Build common_titles_map: machinery → set of titles common to both files
            common_titles_map = {}
            if not title_diff_df.empty and 'Common Titles' in title_diff_df.columns:
                for _, row in title_diff_df.iterrows():
                    machinery = row['Machinery']
                    raw = row.get('Common Titles', '-') or '-'
                    if raw.strip() and raw.strip() != '-':
                        common_titles_map[machinery] = {
                            t.strip() for t in raw.split(', ') if t.strip()
                        }
                    else:
                        common_titles_map[machinery] = set()

            # Re-run frequency comparison filtered to common titles only
            freq_df_filtered, freq_excel_filtered = build_frequency_comparison(
                job_detail['detail1'], job_detail['detail2'],
                job_detail['col1'], job_detail['col2'],
                common_titles_map=common_titles_map
            )
            job_detail['freq_df']    = freq_df_filtered
            job_detail['freq_excel'] = freq_excel_filtered

            st.session_state.title_diff_df = title_diff_df
            st.session_state.machinery_diff_list = machinery_diff_list
            st.session_state.title_excel_data = title_excel_data
            st.session_state.count_comparison_df = count_comparison_df
            st.session_state.count_excel_data = count_excel_data
            st.session_state.job_detail = job_detail

            st.success("Files processed successfully! View results in the tabs below.")
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        st.exception(e)

tab1, tab2, tab3 = st.tabs(["Job Title Comparison", "Machinery Count Comparison", "Frequency Interval Comparison"])

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

            title_cols_to_format = [c for c in diff_only_df.columns if 'Titles only in' in c or 'Common Titles' in c]
            for col in title_cols_to_format:
                def format_with_count(row, c=col):
                    val = row[c]
                    if val == '-' or pd.isna(val):
                        return val
                    count = len([x for x in str(val).split(', ') if x.strip()])
                    return f"{val}\n(count: {count})" if count > 0 else val
                diff_only_df[col] = diff_only_df.apply(format_with_count, axis=1)

            def highlight_title_counts(row):
                styles = [''] * len(row)
                for idx, c in enumerate(row.index):
                    if 'Titles only in' in c and row[c] != '-':
                        styles[idx] = 'background-color: #FFF3E0'
                    elif 'Common Titles' in c and row[c] != '-':
                        styles[idx] = 'background-color: #E8F5E9'
                    elif 'Count' in c:
                        try:
                            if isinstance(row[c], (int, float)) and row[c] > 0:
                                styles[idx] = 'background-color: #E3E1F7'
                        except Exception:
                            pass
                return styles

            styled_df = diff_only_df.style.apply(highlight_title_counts, axis=1)
            st.dataframe(styled_df, use_container_width=True)

            if isinstance(title_excel_data, bytes) and len(title_excel_data) > 0:
                st.subheader("📥 Download Report")
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
        job_detail = st.session_state.job_detail

        col1_name = comparison_df.columns[1]
        col2_name = comparison_df.columns[2]

        def highlight_differences(row):
            styles = [''] * len(row)
            if row['Machinery'] != 'TOTAL':
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
            label="📥 Download Excel Report (Count Summary + Job Detail Breakdown)",
            data=excel_data,
            file_name="Machinery_Count_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.info("""
        **Legend:**
        - **Red highlighting**: Machinery that only exists in one file
        - **Yellow highlighting**: Different job counts between files
        - **Green (positive difference)**: More jobs in first file
        - **Red (negative difference)**: More jobs in second file
        """)

        # --- Job Title & Code breakdown for machinery with differences ---
        if job_detail:
            diff_rows = comparison_df[
                (comparison_df['Machinery'] != 'TOTAL') &
                (comparison_df[col1_name] != comparison_df[col2_name])
            ]['Machinery'].tolist()

            if diff_rows:
                st.markdown("---")
                st.subheader("🔍 Job Title & Code Breakdown for Differing Machinery")
                st.write(
                    f"Showing **{len(diff_rows)}** machinery items where job counts differ. "
                    "Duplicate Job Codes are highlighted with their occurrence count."
                )

                detail1 = job_detail.get('detail1', {})
                detail2 = job_detail.get('detail2', {})
                label1 = job_detail.get('col1', 'File 1')
                label2 = job_detail.get('col2', 'File 2')

                for machinery in diff_rows:
                    with st.expander(f"📋 {machinery}", expanded=False):
                        c1, c2 = st.columns(2)

                        def render_detail_table(detail_dict, machinery_name, label):
                            df_detail = detail_dict.get(machinery_name, pd.DataFrame(
                                columns=['Job Code', 'Job Title', 'Count']
                            ))
                            if df_detail.empty:
                                st.write(f"**{label}**")
                                st.info("No jobs found in this file.")
                                return

                            st.write(f"**{label}** — {int(df_detail['Count'].sum())} job(s)")

                            # Highlight rows where Job Code appears more than once
                            dup_codes = set(
                                df_detail.loc[df_detail['Count'] > 1, 'Job Code'].tolist()
                            )

                            def highlight_dups(row):
                                if row['Count'] > 1:
                                    return ['background-color: #FFF3CD'] * len(row)
                                return [''] * len(row)

                            if not df_detail.empty:
                                styled = df_detail.style.apply(highlight_dups, axis=1)
                                st.dataframe(styled, use_container_width=True, hide_index=True)
                                if dup_codes:
                                    st.caption(
                                        f"⚠️ Duplicate Job Codes: {', '.join(sorted(dup_codes))}"
                                    )

                        with c1:
                            render_detail_table(detail1, machinery, label1)
                        with c2:
                            render_detail_table(detail2, machinery, label2)
    else:
        st.info("Please upload both CSV files to generate the machinery count comparison report.")

with tab3:
    st.header("Frequency Interval Comparison")
    st.markdown(
        "For each job code, compares the **Frequency interval** (e.g. *3 Months*, *12 Months*) "
        "between the two files — **limited to job titles that are common to both files**. "
        "Machinery with no shared titles is excluded."
    )
    job_detail = st.session_state.job_detail
    if job_detail is not None:
        freq_df    = job_detail.get('freq_df')
        freq_excel = job_detail.get('freq_excel')
        label1     = job_detail.get('col1', 'File 1')
        label2     = job_detail.get('col2', 'File 2')

        if freq_df is not None and not freq_df.empty and 'Match' in freq_df.columns:
            total_codes  = len(freq_df)
            matched      = len(freq_df[freq_df['Match'] == '✓ Match'])
            differ       = len(freq_df[freq_df['Match'] == '✗ Differ'])
            only_in_f1   = len(freq_df[freq_df['Match'] == 'Only in File 1'])
            only_in_f2   = len(freq_df[freq_df['Match'] == 'Only in File 2'])

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total Job Codes", total_codes)
            m2.metric("✓ Same Frequency", matched)
            m3.metric("✗ Frequency Differs", differ)
            m4.metric("Only in File 1", only_in_f1)
            m5.metric("Only in File 2", only_in_f2)

            st.info("""
**Legend:**
- 🟢 **Green row**: Same frequency interval in both files
- 🔴 **Red row**: Frequency interval differs between files
- 🟠 **Orange row**: Job code only exists in File 1 (for a shared title)
- 🔵 **Blue row**: Job code only exists in File 2 (for a shared title)
""")

            st.download_button(
                label="📥 Download Frequency Interval Comparison Excel",
                data=freq_excel,
                file_name="Job_Frequency_Interval_Comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.markdown("---")

            def highlight_freq_row(row):
                match = row.get('Match', '')
                n = len(row)
                if match == '✓ Match':
                    return ['background-color: #C6EFCE; color: #006100'] * n
                elif match == '✗ Differ':
                    return ['background-color: #FFC7CE; color: #9C0006; font-weight: bold'] * n
                elif match == 'Only in File 1':
                    return ['background-color: #FCE4D6; color: #974706'] * n
                elif match == 'Only in File 2':
                    return ['background-color: #BDD7EE; color: #0070C0'] * n
                return [''] * n

            all_machinery = freq_df['Machinery'].unique().tolist()
            col1_name = label1
            col2_name = label2

            for machinery in all_machinery:
                mach_df = freq_df[freq_df['Machinery'] == machinery].drop(columns=['Machinery']).reset_index(drop=True)
                has_diff = mach_df['Match'].isin(['✗ Differ', 'Only in File 1', 'Only in File 2']).any()
                icon = "⚠️" if has_diff else "✅"
                diff_count = mach_df['Match'].isin(['✗ Differ', 'Only in File 1', 'Only in File 2']).sum()
                label_text = f"{icon} {machinery}  —  {diff_count} difference(s) in frequency interval"

                with st.expander(label_text, expanded=False):
                    styled = mach_df.style.apply(highlight_freq_row, axis=1)
                    st.dataframe(styled, use_container_width=True, hide_index=True)
        else:
            st.info("No frequency interval data available. Please re-upload both CSV files.")
    else:
        st.info("Please upload both CSV files to generate the frequency interval comparison report.")
