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

# File uploaders in the main area (outside tabs)
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
- Supports comparing files from the same vessel (adds file identifiers to columns)
- Improved visual organization with expandable sections
""")

# File uploaders
col1, col2 = st.columns(2)

with col1:
    st.subheader("First File")
    file1 = st.file_uploader("Upload first CSV file", type=["csv"])
    
with col2:
    st.subheader("Second File")
    file2 = st.file_uploader("Upload second CSV file", type=["csv"])

# Create session state for storing results
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

# Process files when both are uploaded
if file1 is not None and file2 is not None:
    try:
        file1_content = file1.getvalue()
        file2_content = file2.getvalue()
        
        with st.spinner("Processing files for both comparisons..."):
            # Process title comparison
            title_diff_df, machinery_diff_list, title_excel_data = compare_titles(
                file1_content, file2_content, file1.name, file2.name
            )
            
            # Process count comparison
            count_comparison_df, count_excel_data = process_files(
                file1_content, file2_content, file1.name, file2.name
            )
            
            # Store results in session state
            st.session_state.title_diff_df = title_diff_df
            st.session_state.machinery_diff_list = machinery_diff_list
            st.session_state.title_excel_data = title_excel_data
            st.session_state.count_comparison_df = count_comparison_df
            st.session_state.count_excel_data = count_excel_data
            
            st.success("Files processed successfully! View results in the tabs below.")
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        st.exception(e)

# Create tabs for different views
tab1, tab2 = st.tabs(["Job Title Comparison", "Machinery Count Comparison"])

# Job Title Comparison Tab
with tab1:
    st.header("Job Title Comparison Results")
    
    if st.session_state.title_diff_df is not None and st.session_state.machinery_diff_list is not None:
        title_diff_df = st.session_state.title_diff_df
        machinery_diff_list = st.session_state.machinery_diff_list
        title_excel_data = st.session_state.title_excel_data
        
        # Display summary statistics
        st.subheader("📊 Comparison Summary")
        total_machinery = len(title_diff_df) if isinstance(title_diff_df, pd.DataFrame) else 0
        
        if isinstance(title_diff_df, pd.DataFrame) and not title_diff_df.empty:
            diff_count = len(machinery_diff_list)
            same_count = total_machinery - diff_count
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Machinery Items", total_machinery)
            with col2:
                st.metric("Items with Different Titles", diff_count)
            with col3:
                st.metric("Items with Same Titles", same_count)
            
            # Only continue if we have differences
            if diff_count > 0:
                # Main results section
                st.subheader("📋 Machinery with Different Job Titles")
                st.write(f"There are **{diff_count}** machinery items with different job titles:")
                
                # Display a list of machinery with differences 
                machinery_list_text = "\n".join([f"• {machinery}" for machinery in machinery_diff_list])
                st.text_area("Machinery List:", machinery_list_text, height=150)
                
                # Display the comparison table
                st.subheader("🔄 Detailed Title Comparison")
                
                # Filter to only show rows with differences for clarity
                diff_only_df = title_diff_df[title_diff_df['Has Differences'] == 'Yes'].copy()
                
                # Add index for easier reference
                diff_only_df = diff_only_df.reset_index(drop=True)
                diff_only_df.index = diff_only_df.index + 1  # Start from 1 instead of 0
                
                # Display the table with differences
                st.dataframe(diff_only_df, use_container_width=True)
                
                # Show raw data for inspection
            #     st.subheader("🔎 Examples of Job Title Differences")
                
            #     # Sample up to 5 machinery items to show detailed differences
            #     sample_count = min(5, len(diff_only_df))
            #     if sample_count > 0:
            #         st.write("Below are examples of machinery with different job titles:")
                    
            #         sample_machines = diff_only_df['Machinery'].head(sample_count).tolist()
                    
            #         for idx, machinery in enumerate(sample_machines):
            #             row = diff_only_df[diff_only_df['Machinery'] == machinery].iloc[0]
                        
            #             st.write(f"**{idx+1}. {machinery}**")
                        
            #             # Use expander for better organization of content
            #             with st.expander(f"View all title details for {machinery}", expanded=True):
            #                 # Get the title columns
            #                 title_cols = [col for col in diff_only_df.columns if col.startswith('Titles only in')]
                            
            #                 # Display common titles first (if any)
            #                 st.write("**Common Titles:**")
            #                 if 'Common Titles' in row and row['Common Titles'] != '-':
            #                     st.markdown(
            #                         f"<div style='background-color: #E8F5E9; padding: 10px; border-radius: 5px;'>{row['Common Titles']}</div>", 
            #                         unsafe_allow_html=True
            #                     )
            #                 else:
            #                     st.write("*None*")
                            
            #                 st.markdown("---")
                            
            #                 # Display titles from both files in separate sections
            #                 cols = st.columns(2)
                            
            #                 # First file titles
            #                 with cols[0]:
            #                     if len(title_cols) > 0:
            #                         first_col = title_cols[0]
            #                         st.write(f"**{first_col}:**")
            #                         if row[first_col] != '-':
            #                             st.markdown(
            #                                 f"<div style='background-color: #FFF3E0; padding: 10px; border-radius: 5px;'>{row[first_col]}</div>", 
            #                                 unsafe_allow_html=True
            #                             )
            #                         else:
            #                             st.write("*None*")
                            
            #                 # Second file titles
            #                 with cols[1]:
            #                     if len(title_cols) > 1:
            #                         second_col = title_cols[1]
            #                         st.write(f"**{second_col}:**")
            #                         if row[second_col] != '-':
            #                             st.markdown(
            #                                 f"<div style='background-color: #E3F2FD; padding: 10px; border-radius: 5px;'>{row[second_col]}</div>", 
            #                                 unsafe_allow_html=True
            #                             )
            #                         else:
            #                             st.write("*None*")
                        
            #             st.write("---")
            # else:
            #     st.success("No job title differences found for any machinery!")
                
            # Download section
            if isinstance(title_excel_data, bytes) and len(title_excel_data) > 0:
                st.subheader("📥 Download Report")
                st.write("Download the detailed Excel report with highlighted job title differences:")
                st.download_button(
                    label="Download Job Title Comparison Report",
                    data=title_excel_data,
                    file_name="Job_Title_Comparison.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Click to download the job title comparison report in Excel format"
                )
        else:
            st.info("No job title comparison data generated. Please check if both files have matching machinery.")
    else:
        st.info("Please upload both CSV files to generate the title comparison report.")

# Machinery Count Comparison Tab
with tab2:
    st.header("Machinery Count Comparison Results")
    
    if st.session_state.count_comparison_df is not None:
        comparison_df = st.session_state.count_comparison_df
        excel_data = st.session_state.count_excel_data
        
        # Create a styled dataframe with custom formatting for display
        def highlight_differences(row):
            styles = [''] * len(row)
            
            if row['Machinery'] != 'TOTAL':
                col1_name = comparison_df.columns[1]  # First vessel column
                col2_name = comparison_df.columns[2]  # Second vessel column
                
                # Missing in one file (red)
                if row[col1_name] == 0 or row[col2_name] == 0:
                    styles[0] = 'background-color: #FFC7CE; font-weight: bold'  # Red for machinery
                    styles[3] = 'background-color: #FFC7CE; color: #9C0006'     # Red for diff
                
                # Different counts (yellow for counts)
                if row[col1_name] != row[col2_name]:
                    styles[1] = 'background-color: #FFEB9C'  # Yellow for first count
                    styles[2] = 'background-color: #FFEB9C'  # Yellow for second count
                    
                    # More in first file (green for positive diff)
                    if row[col1_name] > row[col2_name]:
                        styles[3] = 'background-color: #C6EFCE; color: #006100'  # Green for diff
                    # More in second file (red for negative diff)
                    else:
                        styles[3] = 'background-color: #FFC7CE; color: #9C0006'  # Red for diff
            else:
                # Total row (bold)
                return ['font-weight: bold'] * len(row)
            
            return styles
        
        # Apply styling and display
        styled_df = comparison_df.style.apply(highlight_differences, axis=1)
        st.dataframe(styled_df, use_container_width=True)
        
        # Download button for Excel file
        st.download_button(
            label="Download Excel Report",
            data=excel_data,
            file_name="Machinery_Count_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Show explanation
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