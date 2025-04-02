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
    if "Vessel" in df.columns:
        vessel = df["Vessel"].dropna().astype(str).iloc[0].strip()
        return vessel
    return "Unknown Vessel"


def rename_machinery(value):
    """Apply renaming rules to machinery values."""
    rename_mapping = {
        r"P1$": " P", r"Port1$": " P", r"S1$": " S", r"Starboard1$": " S", 
        r"S2$": " S", r"Starboard2$": " S", r"F$": " F", r"Forward$": " F", 
        r"S1$": " S", r"Starboard1$": " S", r"F1$": " F", r"Forward1$": " F",
        r"S2$": " S", r"Starboard2$": " S", r"F2$": " F", r"Forward2$": " F",
        r"A$": " A", r"Aft$": " A", r"P$": " P", r"Port$": " P", r"S$": " S",
        r"A1$": " A", r"Aft1$": " A", r"P1$": " P", r"Port1$": " P", r"S$": " S",
        r"A2$": " A", r"Aft2$": " A", r"P2$": " P", r"Port2$": " P", r"S$": " S",
        r"Aft- P$": " A-P", r"A- P$": " A-P", r"F- P$": " F-P", r"Fwd- P$": " F-P", r"F- S$": " F-S",r"Fwd-Stbd$": " F-S",
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
        r"LiferaftS$": " Liferaft S", r"LiferaftStarboard$": " Liferaft S",
        r"Provision CraneA- P$": "Provision Crane A-P",
        r"Provision CraneAft- P$": "Provision Crane A-P",
        r"Provision CraneF- P$": "Provision Crane F-P",
        r"Provision CraneF- S$": "Provision Crane F-S",
        r"Provision CraneFwd- P$": "Provision Crane F-P"
    }
    
    original_value = str(value).strip()
    for pattern, replacement in rename_mapping.items():
        if re.search(pattern, original_value):
            return re.sub(pattern, replacement, original_value)
    return original_value

def process_files(file1_content, file2_content, file1_name, file2_name):
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font
    import os
    from io import BytesIO

    def extract_date_from_filename(filename):
        base_name = os.path.splitext(os.path.basename(filename))[0]
        date_part = base_name.split()[-1]
        return f"{date_part[0:2]}-{date_part[2:4]}-{date_part[4:]}"

    def get_vessel_name(df):
        if "Vessel" in df.columns:
            return str(df["Vessel"].dropna().iloc[0]).strip()
        return "Unknown Vessel"

    df_system_mgmt = pd.read_csv(BytesIO(file1_content))
    df_pms_jobs = pd.read_csv(BytesIO(file2_content))

    date1_fmt = extract_date_from_filename(file1_name)
    date2_fmt = extract_date_from_filename(file2_name)
    vessel1 = get_vessel_name(df_system_mgmt)
    vessel2 = get_vessel_name(df_pms_jobs)

    col1 = f"{vessel1} ({date1_fmt})"
    col2 = f"{vessel2} ({date2_fmt})"

    # Ensure uniqueness if vessel names are the same
    if col1 == col2:
        col1 += " [File 1]"
        col2 += " [File 2]"

    print("[DEBUG] col1:", repr(col1))
    print("[DEBUG] col2:", repr(col2))

    # Auto-detect machinery column
    possible_machinery_columns = ['Machinery', 'Machinery Location', 'Component Name', 'System Name']

    for col in possible_machinery_columns:
        if col in df_system_mgmt.columns:
            df_system_mgmt.rename(columns={col: 'Machinery'}, inplace=True)
            break
    else:
        raise ValueError("No recognized Machinery column in first file.")

    for col in possible_machinery_columns:
        if col in df_pms_jobs.columns:
            df_pms_jobs.rename(columns={col: 'Machinery Location'}, inplace=True)
            break
    else:
        raise ValueError("No recognized Machinery column in second file.")

    from comparison_utils import rename_machinery  # use your existing function

    df_system_mgmt['Machinery'] = df_system_mgmt['Machinery'].apply(rename_machinery)
    df_pms_jobs['Machinery Location'] = df_pms_jobs['Machinery Location'].apply(rename_machinery)

    system_mgmt_counts = df_system_mgmt['Machinery'].value_counts().reset_index()
    pms_jobs_counts = df_pms_jobs['Machinery Location'].value_counts().reset_index()

    if system_mgmt_counts.shape[1] != 2:
        raise ValueError("Unexpected structure in system_mgmt_counts:\n" + str(system_mgmt_counts.head()))
    if pms_jobs_counts.shape[1] != 2:
        raise ValueError("Unexpected structure in pms_jobs_counts:\n" + str(pms_jobs_counts.head()))

    system_mgmt_counts.columns = ['Machinery', col1]
    pms_jobs_counts.columns = ['Machinery', col2]

    comparison_df = pd.merge(system_mgmt_counts, pms_jobs_counts, on='Machinery', how='outer').fillna(0)

    print("[DEBUG] Actual merged DataFrame columns:", comparison_df.columns.tolist())

    if col1 not in comparison_df.columns or col2 not in comparison_df.columns:
        raise KeyError(
            f"Column mismatch!\nExpected: {col1}, {col2}\nActual: {comparison_df.columns.tolist()}"
        )

    comparison_df[col1] = comparison_df[col1].astype(int)
    comparison_df[col2] = comparison_df[col2].astype(int)
    comparison_df['Difference'] = comparison_df[col1] - comparison_df[col2]

    total_row = {
        'Machinery': 'TOTAL',
        col1: comparison_df[col1].sum(),
        col2: comparison_df[col2].sum(),
        'Difference': comparison_df["Difference"].sum()
    }
    comparison_df = pd.concat([comparison_df, pd.DataFrame([total_row])], ignore_index=True)

    # Excel generation
    output = BytesIO()
    comparison_df.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    sheet = wb.active

    # Styles
    fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    bold_font = Font(bold=True)
    red_font = Font(color="9C0006")
    green_font = Font(color="006100")

    # Highlighting
    for row in range(2, sheet.max_row + 1):
        machinery = sheet.cell(row=row, column=1)
        count1 = sheet.cell(row=row, column=2).value
        count2 = sheet.cell(row=row, column=3).value
        diff_cell = sheet.cell(row=row, column=4)

        if machinery.value != 'TOTAL':
            if count1 == 0 or count2 == 0:
                machinery.fill = fill_red
                machinery.font = bold_font
                diff_cell.fill = fill_red
                diff_cell.font = red_font
            if count1 != count2:
                sheet.cell(row=row, column=2).fill = fill_yellow
                sheet.cell(row=row, column=3).fill = fill_yellow
                if count1 > count2:
                    diff_cell.fill = fill_green
                    diff_cell.font = green_font
                else:
                    diff_cell.fill = fill_red
                    diff_cell.font = red_font
        else:
            for col in range(1, 5):
                sheet.cell(row=row, column=col).font = bold_font

    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)

    return comparison_df, output_final.getvalue()


