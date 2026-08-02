import streamlit as st
import pandas as pd
from new_title_comparison import compare_titles
from comparison_utils import process_files, build_frequency_comparison, build_new_jobs_analysis, build_critical_jobs_analysis
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
if 'new_jobs_analysis' not in st.session_state:
    st.session_state.new_jobs_analysis = None
if 'critical_jobs_analysis' not in st.session_state:
    st.session_state.critical_jobs_analysis = None

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

            try:
                new_summary, new_detail, new_excel, new_label1, new_label2 = build_new_jobs_analysis(
                    file1_content, file2_content, file1.name, file2.name
                )
                st.session_state.new_jobs_analysis = {
                    'summary': new_summary,
                    'detail': new_detail,
                    'excel': new_excel,
                    'label1': new_label1,
                    'label2': new_label2,
                }
            except Exception as nj_err:
                st.session_state.new_jobs_analysis = {'error': str(nj_err)}

            try:
                crit_summary, crit_detail, crit_excel, crit_label1, crit_label2 = build_critical_jobs_analysis(
                    file1_content, file2_content, file1.name, file2.name
                )
                st.session_state.critical_jobs_analysis = {
                    'summary': crit_summary,
                    'detail': crit_detail,
                    'excel': crit_excel,
                    'label1': crit_label1,
                    'label2': crit_label2,
                }
            except Exception as cj_err:
                st.session_state.critical_jobs_analysis = {'error': str(cj_err)}

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

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Job Title Comparison", "Machinery Count Comparison", "Frequency Interval Comparison", "New Jobs Analysis", "Critical Jobs"])

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

        st.markdown("""
<div style="border:1px solid #ddd; border-radius:6px; padding:12px 16px; background:#fafafa; margin-top:8px;">
<strong>Legend</strong><br><br>
<span style="display:inline-block;width:18px;height:18px;background:#FFC7CE;border:1px solid #ccc;vertical-align:middle;margin-right:6px;"></span> Machinery that only exists in one file<br><br>
<span style="display:inline-block;width:18px;height:18px;background:#FFEB9C;border:1px solid #ccc;vertical-align:middle;margin-right:6px;"></span> Different job counts between files<br><br>
<span style="display:inline-block;width:18px;height:18px;background:#C6EFCE;border:1px solid #ccc;vertical-align:middle;margin-right:6px;"></span> Difference is positive — more jobs in first file<br><br>
<span style="display:inline-block;width:18px;height:18px;background:#FFC7CE;border:1px solid #ccc;vertical-align:middle;margin-right:6px;"></span> Difference is negative — more jobs in second file
</div>
""", unsafe_allow_html=True)

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

                st.markdown("""
<div style="border:1px solid #ddd;border-radius:6px;padding:10px 14px;background:#fafafa;margin-bottom:10px;font-size:0.9em;">
<strong>Breakdown Legend</strong>&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#FFD180;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Only in left file&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#BBDEFB;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Only in right file&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#FFF3CD;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Duplicate Job Code
</div>
""", unsafe_allow_html=True)

                for machinery in diff_rows:
                    with st.expander(f"📋 {machinery}", expanded=False):
                        c1, c2 = st.columns(2)

                        df1_m = detail1.get(machinery, pd.DataFrame(
                            columns=['Job Code', 'Job Title', 'Count']))
                        df2_m = detail2.get(machinery, pd.DataFrame(
                            columns=['Job Code', 'Job Title', 'Count']))

                        codes1 = set(df1_m['Job Code'].astype(str).str.strip()) if not df1_m.empty else set()
                        codes2 = set(df2_m['Job Code'].astype(str).str.strip()) if not df2_m.empty else set()
                        only_in_1 = codes1 - codes2
                        only_in_2 = codes2 - codes1

                        def render_detail_table(df_detail, label, exclusive_codes, exclusive_color):
                            if df_detail.empty:
                                st.write(f"**{label}**")
                                st.info("No jobs found in this file.")
                                return
                            st.write(f"**{label}** — {int(df_detail['Count'].sum())} job(s)")

                            def highlight_row(row, _exc=exclusive_codes, _col=exclusive_color):
                                code = str(row.get('Job Code', '')).strip()
                                if code in _exc:
                                    return [f'background-color: {_col}'] * len(row)
                                if row.get('Count', 1) > 1:
                                    return ['background-color: #FFF3CD'] * len(row)
                                return [''] * len(row)

                            styled = df_detail.style.apply(highlight_row, axis=1)
                            st.dataframe(styled, use_container_width=True, hide_index=True)

                            dup_codes = set(df_detail.loc[df_detail['Count'] > 1, 'Job Code'].tolist())
                            if dup_codes:
                                st.caption(f"⚠️ Duplicate Job Codes: {', '.join(sorted(map(str, dup_codes)))}")

                        with c1:
                            render_detail_table(df1_m, label1, only_in_1, '#FFD180')
                        with c2:
                            render_detail_table(df2_m, label2, only_in_2, '#BBDEFB')
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

with tab4:
    st.header("New Jobs Analysis")
    st.markdown(
        "Filters **both files** to rows where **Job Status = New**, "
        "then compares those jobs per machinery location."
    )

    nja = st.session_state.new_jobs_analysis

    if nja is None:
        st.info("Please upload both CSV files to generate the New Jobs analysis.")
    elif 'error' in nja:
        st.warning(f"New Jobs Analysis could not run: {nja['error']}")
    else:
        summary_df = nja['summary']
        detail     = nja['detail']
        new_excel  = nja['excel']
        lbl1       = nja['label1']
        lbl2       = nja['label2']

        col_new1   = f'New in {lbl1}'
        col_new2   = f'New in {lbl2}'
        col_common = 'Common in Both'
        col_not1   = f'Not in {lbl1}'
        col_not2   = f'Not in {lbl2}'

        # ── Summary metrics ──────────────────────────────────────────────────
        total_mach   = len(summary_df)
        total_new1   = int(summary_df[col_new1].sum())   if col_new1   in summary_df.columns else 0
        total_new2   = int(summary_df[col_new2].sum())   if col_new2   in summary_df.columns else 0
        total_common = int(summary_df[col_common].sum()) if col_common in summary_df.columns else 0
        total_not1   = int(summary_df[col_not1].sum())   if col_not1   in summary_df.columns else 0
        total_not2   = int(summary_df[col_not2].sum())   if col_not2   in summary_df.columns else 0

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Machinery Items",       total_mach)
        c2.metric(f"New in {lbl1}",        total_new1)
        c3.metric(f"New in {lbl2}",        total_new2)
        c4.metric("✅ Common in Both",     total_common)
        c5.metric(f"🟠 Not in {lbl1}",    total_not1)
        c6.metric(f"🔵 Not in {lbl2}",    total_not2)

        # ── Legend ───────────────────────────────────────────────────────────
        st.markdown(f"""
<div style="border:1px solid #ddd;border-radius:6px;padding:10px 14px;background:#fafafa;margin:8px 0;font-size:0.9em;">
<strong>Legend</strong>&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#C6EFCE;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Common in Both — job code is New in both files&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#FFD180;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Not in {lbl1} — New in {lbl2} but absent from {lbl1}&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#BBDEFB;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Not in {lbl2} — New in {lbl1} but absent from {lbl2}
</div>
""", unsafe_allow_html=True)

        # ── Summary table ─────────────────────────────────────────────────────
        st.subheader("📊 Summary by Machinery Location")

        def highlight_summary(row):
            styles = [''] * len(row)
            cols = list(row.index)
            if col_common in cols and row[col_common] > 0:
                styles[cols.index(col_common)] = 'background-color: #C6EFCE'
            if col_not1 in cols and row[col_not1] > 0:
                styles[cols.index(col_not1)] = 'background-color: #FFD180'
            if col_not2 in cols and row[col_not2] > 0:
                styles[cols.index(col_not2)] = 'background-color: #BBDEFB'
            return styles

        styled_sum = summary_df.style.apply(highlight_summary, axis=1)
        st.dataframe(styled_sum, use_container_width=True, hide_index=True)

        st.download_button(
            label="📥 Download New Jobs Analysis Excel",
            data=new_excel,
            file_name="New_Jobs_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ── Per-machinery detail expanders ────────────────────────────────────
        if detail:
            st.markdown("---")
            st.subheader("🔍 Job Detail by Machinery Location")

            for machinery, df_det in detail.items():
                n_both = len(df_det[df_det['Match'] == 'In Both'])
                n_not1 = len(df_det[df_det['Match'] == f'Only in {lbl1}'])
                n_not2 = len(df_det[df_det['Match'] == f'Only in {lbl2}'])
                has_diff = n_not1 > 0 or n_not2 > 0
                icon = "⚠️" if has_diff else "✅"
                label_text = (
                    f"{icon} {machinery}  —  "
                    f"{n_both} common · {n_not1} only in {lbl1} · {n_not2} only in {lbl2}"
                )

                with st.expander(label_text, expanded=False):
                    def highlight_det(row, _l1=lbl1, _l2=lbl2):
                        m = row.get('Match', '')
                        if m == 'In Both':
                            return ['background-color: #C6EFCE; color: #006100'] * len(row)
                        elif m == f'Only in {_l1}':
                            return ['background-color: #BBDEFB; color: #004494'] * len(row)
                        elif m == f'Only in {_l2}':
                            return ['background-color: #FFD180; color: #7D4E00'] * len(row)
                        return [''] * len(row)

                    styled_det = df_det.style.apply(highlight_det, axis=1)
                    st.dataframe(styled_det, use_container_width=True, hide_index=True)
        else:
            st.info("No 'New' status jobs found in either file.")

with tab5:
    st.header("Critical Jobs Comparison")
    st.markdown(
        "Filters **both files** to rows where **Column B = C** (critical jobs), "
        "then compares those jobs per machinery location between the two files."
    )

    cja = st.session_state.critical_jobs_analysis

    if cja is None:
        st.info("Please upload both CSV files to generate the Critical Jobs analysis.")
    elif 'error' in cja:
        st.warning(f"Critical Jobs Analysis could not run: {cja['error']}")
    else:
        summary_df = cja['summary']
        detail     = cja['detail']
        crit_excel = cja['excel']
        lbl1       = cja['label1']
        lbl2       = cja['label2']

        col_crit1  = f'Critical in {lbl1}'
        col_crit2  = f'Critical in {lbl2}'
        col_common = 'Common in Both'
        col_not1   = f'Not in {lbl1}'
        col_not2   = f'Not in {lbl2}'

        # ── Summary metrics ──────────────────────────────────────────────────
        total_mach    = len(summary_df)
        total_crit1   = int(summary_df[col_crit1].sum())  if col_crit1  in summary_df.columns else 0
        total_crit2   = int(summary_df[col_crit2].sum())  if col_crit2  in summary_df.columns else 0
        total_common  = int(summary_df[col_common].sum()) if col_common in summary_df.columns else 0
        total_not1    = int(summary_df[col_not1].sum())   if col_not1   in summary_df.columns else 0
        total_not2    = int(summary_df[col_not2].sum())   if col_not2   in summary_df.columns else 0

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Machinery Items",        total_mach)
        c2.metric(f"Critical in {lbl1}",    total_crit1)
        c3.metric(f"Critical in {lbl2}",    total_crit2)
        c4.metric("✅ Common in Both",      total_common)
        c5.metric(f"🟠 Not in {lbl1}",     total_not1)
        c6.metric(f"🔵 Not in {lbl2}",     total_not2)

        # ── Legend ───────────────────────────────────────────────────────────
        st.markdown(f"""
<div style="border:1px solid #ddd;border-radius:6px;padding:10px 14px;background:#fafafa;margin:8px 0;font-size:0.9em;">
<strong>Legend</strong>&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#C6EFCE;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Common in Both — critical job exists in both files&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#FFD180;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Not in {lbl1} — critical in {lbl2} but absent from {lbl1}&nbsp;&nbsp;
<span style="display:inline-block;width:14px;height:14px;background:#BBDEFB;border:1px solid #ccc;vertical-align:middle;margin-right:4px;"></span>Not in {lbl2} — critical in {lbl1} but absent from {lbl2}
</div>
""", unsafe_allow_html=True)

        # ── Summary table ─────────────────────────────────────────────────────
        st.subheader("📊 Summary by Machinery Location")

        def highlight_crit_summary(row):
            styles = [''] * len(row)
            cols = list(row.index)
            if col_common in cols and row[col_common] > 0:
                styles[cols.index(col_common)] = 'background-color: #C6EFCE'
            if col_not1 in cols and row[col_not1] > 0:
                styles[cols.index(col_not1)] = 'background-color: #FFD180'
            if col_not2 in cols and row[col_not2] > 0:
                styles[cols.index(col_not2)] = 'background-color: #BBDEFB'
            return styles

        styled_sum = summary_df.style.apply(highlight_crit_summary, axis=1)
        st.dataframe(styled_sum, use_container_width=True, hide_index=True)

        st.download_button(
            label="📥 Download Critical Jobs Excel",
            data=crit_excel,
            file_name="Critical_Jobs_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ── Per-machinery detail expanders ────────────────────────────────────
        if detail:
            st.markdown("---")
            st.subheader("🔍 Critical Job Detail by Machinery Location")

            for machinery, df_det in detail.items():
                n_both = len(df_det[df_det['Match'] == 'In Both'])
                n_only1 = len(df_det[df_det['Match'] == f'Only in {lbl1}'])
                n_only2 = len(df_det[df_det['Match'] == f'Only in {lbl2}'])
                has_diff = n_only1 > 0 or n_only2 > 0
                icon = "⚠️" if has_diff else "✅"
                label_text = (
                    f"{icon} {machinery}  —  "
                    f"{n_both} common · {n_only1} only in {lbl1} · {n_only2} only in {lbl2}"
                )

                with st.expander(label_text, expanded=False):
                    def highlight_crit_det(row, _l1=lbl1, _l2=lbl2):
                        m = row.get('Match', '')
                        if m == 'In Both':
                            return ['background-color: #C6EFCE; color: #006100'] * len(row)
                        elif m == f'Only in {_l1}':
                            return ['background-color: #BBDEFB; color: #004494'] * len(row)
                        elif m == f'Only in {_l2}':
                            return ['background-color: #FFD180; color: #7D4E00'] * len(row)
                        return [''] * len(row)

                    styled_det = df_det.style.apply(highlight_crit_det, axis=1)
                    st.dataframe(styled_det, use_container_width=True, hide_index=True)
        else:
            st.info("No critical jobs (Column B = 'C') found in either file.")
