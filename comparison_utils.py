import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Color
import re
import os
from io import BytesIO

def extract_date_from_filename(filename):
    """Extract and format date from filename."""
    base_name = os.path.splitext(os.path.basename(filename))[0]
    date_part = base_name.split()[-1]
    return f"{date_part[0:2]}-{date_part[2:4]}-{date_part[4:]}"

def get_vessel_name(df):
    """Extract vessel name from DataFrame."""
    return df["Vessel"].iloc[0] if "Vessel" in df.columns else "Unknown Vessel"

def rename_machinery(value):
    """Apply renaming rules to machinery values."""
    rename_mapping = {
        r"P1$": " P", r"Port1$": " P", r"S1$": " S", r"Starboard1$": " S", 
        r"S2$": " S", r"Starboard2$": " S", r"F$": " F", r"Forward$": " F", 
        r"A$": " A", r"Aft$": " A", r"P$": " P", r"Port$": " P", r"S$": " S", 
        r"Starboard$": " S", r"Lifeboat DavitA$": " Lifeboat Davit A",
        r"Lifeboat DavitAft$": " Lifeboat Davit A", r"LifeboatA$": " Lifeboat A",
        r"LifeboatAft$": " Lifeboat A", r"Liferaft 6 PersonF$": " Liferaft 6 Person F",
        r"Liferaft 6 PersonForward$": " Liferaft 6 Person F",
        r"Liferaft Davit LaunchedP$": " Liferaft Davit Launched P",
        r"Liferaft Davit LaunchedPort$": " Liferaft Davit Launched P",
        r"Liferaft Embarkation LadderF$": " Liferaft Embarkation Ladder F",
        r"Liferaft Embarkation LadderForward$": " Liferaft Embarkation Ladder F",
        r"Liferaft Embarkation LadderP$": " Liferaft Embarkation Ladder P",
        r"Liferaft Embarkation LadderPort$": " Liferaft Embarkation Ladder P",
        r"Liferaft Embarkation LadderS$": " Liferaft Embarkation Ladder S",
        r"Liferaft Embarkation LadderStarboard$": " Liferaft Embarkation Ladder S",
        r"Liferaft/Rescue Boat DavitP$": " Liferaft/Rescue Boat Davit P",
        r"Liferaft/Rescue Boat DavitPort$": " Liferaft/Rescue Boat Davit P",
        r"LiferaftS$": " Liferaft S", r"LiferaftStarboard$": " Liferaft S"
    }
    
    original_value = str(value).strip()
    for pattern, replacement in rename_mapping.items():
        if re.search(pattern, original_value):
            return re.sub(pattern, replacement, original_value)
    return original_value

def process_files(file1_content, file2_content, file1_name, file2_name):
    """Process two CSV files and return comparison DataFrame and Excel file."""
    # Read CSV files
    df_system_mgmt = pd.read_csv(BytesIO(file1_content))
    df_pms_jobs = pd.read_csv(BytesIO(file2_content))
    
    # Extract dates and vessel names
    date1_fmt = extract_date_from_filename(file1_name)
    date2_fmt = extract_date_from_filename(file2_name)
    vessel1 = get_vessel_name(df_system_mgmt)
    vessel2 = get_vessel_name(df_pms_jobs)
    
    # Auto-detect and rename Machinery columns
    possible_machinery_columns = ['Machinery', 'Machinery Location', 'Component Name', 'System Name']
    
    for col in possible_machinery_columns:
        if col in df_system_mgmt.columns:
            df_system_mgmt.rename(columns={col: 'Machinery'}, inplace=True)
            break
    else:
        raise ValueError("No recognized Machinery column found in first file")
    
    for col in possible_machinery_columns:
        if col in df_pms_jobs.columns:
            df_pms_jobs.rename(columns={col: 'Machinery Location'}, inplace=True)
            break
    else:
        raise ValueError("No recognized Machinery column found in second file")
    
    # Apply renaming
    df_system_mgmt['Machinery'] = df_system_mgmt['Machinery'].apply(rename_machinery)
    df_pms_jobs['Machinery Location'] = df_pms_jobs['Machinery Location'].apply(rename_machinery)
    
    # Count jobs
    col1 = f"{vessel1} ({date1_fmt})"
    col2 = f"{vessel2} ({date2_fmt})"
    
    system_mgmt_counts = df_system_mgmt['Machinery'].value_counts().reset_index()
    system_mgmt_counts.columns = ['Machinery', col1]
    
    pms_jobs_counts = df_pms_jobs['Machinery Location'].value_counts().reset_index()
    pms_jobs_counts.columns = ['Machinery', col2]
    
    # Merge with outer join
    comparison_df = pd.merge(system_mgmt_counts, pms_jobs_counts, on='Machinery', how='outer').fillna(0)
    comparison_df[col1] = comparison_df[col1].astype(int)
    comparison_df[col2] = comparison_df[col2].astype(int)
    
    # Add Difference column
    comparison_df["Difference"] = comparison_df[col1] - comparison_df[col2]
    
    # Add TOTAL row
    total_row = {
        'Machinery': 'TOTAL',
        col1: comparison_df[col1].sum(),
        col2: comparison_df[col2].sum(),
        'Difference': comparison_df["Difference"].sum()
    }
    comparison_df = pd.concat([comparison_df, pd.DataFrame([total_row])], ignore_index=True)
    
    # Create Excel file in memory
    output = BytesIO()
    comparison_df.to_excel(output, index=False)
    
    # Load workbook for styling
    output.seek(0)
    wb = load_workbook(output)
    sheet = wb.active
    
    # Define styles
    fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Light red
    fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Light green
    fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # Light yellow
    bold_font = Font(bold=True)
    red_font = Font(color="9C0006")  # Dark red
    green_font = Font(color="006100")  # Dark green
    
    # Apply highlighting
    for row in range(2, sheet.max_row + 1):
        machinery = sheet.cell(row=row, column=1)
        count1 = sheet.cell(row=row, column=2).value
        count2 = sheet.cell(row=row, column=3).value
        diff_cell = sheet.cell(row=row, column=4)
        
        if machinery.value != 'TOTAL':
            if count1 == 0 or count2 == 0:  # Missing in one file
                machinery.fill = fill_red
                machinery.font = bold_font
                diff_cell.fill = fill_red
                diff_cell.font = red_font
            
            if count1 != count2:  # Different counts
                sheet.cell(row=row, column=2).fill = fill_yellow
                sheet.cell(row=row, column=3).fill = fill_yellow
                if count1 > count2:  # More in first file
                    diff_cell.fill = fill_green
                    diff_cell.font = green_font
                else:  # More in second file
                    diff_cell.fill = fill_red
                    diff_cell.font = red_font
        else:  # Total row
            for col in range(1, 5):
                sheet.cell(row=row, column=col).font = bold_font
    
    # Save to BytesIO
    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)
    
    return comparison_df, output_final.getvalue()