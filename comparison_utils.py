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
    """Standardize machinery names by applying specific and generic patterns."""
    original_value = str(value).strip()
    original_value = re.sub(r"\s+", " ", original_value)  # normalize spaces
    original_value = re.sub(r"–|—", "-", original_value)  # normalize dashes

    # Priority 1: Specific edge-case replacements
    specific_mapping = {
        # Provision Cranes (existing + new)
        r"^Provision CraneA-?P$": "Provision Crane A-P",
        r"^Provision CraneAft-?Port$": "Provision Crane A-P",
        r"^Provision CraneF-?P$": "Provision Crane F-P",
        r"^Provision CraneF-?S$": "Provision Crane F-S",
        r"^Provision CraneFwd-?P$": "Provision Crane F-P",
        r"^Provision CraneFwd-?Port$": "Provision Crane F-P",
        r"^Provision CraneFwd-?Stbd$": "Provision Crane F-S",
        r"^Provision Crane F-S$": "Provision Crane F-S",
        r"^Provision CraneP1$": "Provision Crane P1",
        r"^Provision CranePort1$": "Provision Crane P1",
        r"^Provision CraneS1$": "Provision Crane S1",
        r"^Provision CraneStarboard1$": "Provision Crane S1",

        # Liferaft/Rescue Davits
        r"^Liferaft/Rescue Boat DavitS$": "Liferaft/Rescue Boat Davit S",
        r"^Liferaft/Rescue Boat DavitStarboard$": "Liferaft/Rescue Boat Davit S",

        # Rescue Boat
        r"^Rescue BoatS$": "Rescue Boat S",
        r"^Rescue BoatStarboard$": "Rescue Boat S",

        # Chain Locker
        r"^Chain LockerP1$": "Chain Locker P1",
        r"^Chain LockerPort1$": "Chain Locker P1",
        r"^Chain LockerS1$": "Chain Locker S1",
        r"^Chain LockerStarboard1$": "Chain Locker S1",

        # Combined Windlass Mooring Winch
        r"^Combined Windlass Mooring WinchF1$": "Combined Windlass Mooring Winch F1",
        r"^Combined Windlass Mooring WinchF2$": "Combined Windlass Mooring Winch F2",
        r"^Combined Windlass Mooring WinchForward1$": "Combined Windlass Mooring Winch F1",
        r"^Combined Windlass Mooring WinchForward2$": "Combined Windlass Mooring Winch F2",

        # Mooring Winch
        r"^Mooring WinchA1$": "Mooring Winch A1",
        r"^Mooring WinchA2$": "Mooring Winch A2",
        r"^Mooring WinchAft1$": "Mooring Winch A1",
        r"^Mooring WinchAft2$": "Mooring Winch A2",

        # Muster Station
        r"^Muster StationA1$": "Muster Station A1",
        r"^Muster StationAft1$": "Muster Station A1",

        # Accommodation Ladder
        r"^Accommodation LadderP1$": "Accommodation Ladder P1",
        r"^Accommodation LadderPort1$": "Accommodation Ladder P1",
        r"^Accommodation LadderS1$": "Accommodation Ladder S1",
        r"^Accommodation LadderStarboard1$": "Accommodation Ladder S1",

        # Anchor Chain Cable
        r"^Anchor Chain CableP1$": "Anchor Chain Cable P1",
        r"^Anchor Chain CablePort1$": "Anchor Chain Cable P1",
        r"^Anchor Chain CableS1$": "Anchor Chain Cable S1",
        r"^Anchor Chain CableStarboard1$": "Anchor Chain Cable S1",

        # Anchor
        r"^AnchorP1$": "Anchor P1",
        r"^AnchorPort1$": "Anchor P1",
        r"^AnchorS1$": "Anchor S1",
        r"^AnchorStarboard1$": "Anchor S1",

        # Pilot Combination Ladder
        # Pilot Combination Ladder
    r"^Pilot Combination LadderP1$": "Pilot Combination Ladder P1",
    r"^Pilot Combination LadderPort1$": "Pilot Combination Ladder P1",
    r"^Pilot Combination LadderS1$": "Pilot Combination Ladder S1",
    r"^Pilot Combination LadderStarboard1$": "Pilot Combination Ladder S1",

    # Bunker Davit
    r"^Bunker DavitP1$": "Bunker Davit P1",
    r"^Bunker DavitPort1$": "Bunker Davit P1",
    r"^Bunker DavitS1$": "Bunker Davit S1",
    r"^Bunker DavitStarboard1$": "Bunker Davit S1",

    # Combined Windlass Mooring Winch
    r"^Combined Windlass Mooring WinchP1$": "Combined Windlass Mooring Winch P1",
    r"^Combined Windlass Mooring WinchPort1$": "Combined Windlass Mooring Winch P1",
    r"^Combined Windlass Mooring WinchS1$": "Combined Windlass Mooring Winch S1",
    r"^Combined Windlass Mooring WinchStarboard1$": "Combined Windlass Mooring Winch S1",

    # Pilot Ladder Davit
    r"^Pilot Ladder DavitP1$": "Pilot Ladder Davit P1",
    r"^Pilot Ladder DavitPort1$": "Pilot Ladder Davit P1",
    r"^Pilot Ladder DavitS2$": "Pilot Ladder Davit S1",
    r"^Pilot Ladder DavitStarboard2$": "Pilot Ladder Davit S1",

    # Seaway Equipment
    r"^Seaway EquipmentP1$": "Seaway Equipment P1",
    r"^Seaway EquipmentPort1$": "Seaway Equipment P1",
    r"^Seaway EquipmentS1$": "Seaway Equipment S1",
    r"^Seaway EquipmentStarboard1$": "Seaway Equipment S1",
    }
    
    for pattern, replacement in specific_mapping.items():
        if re.match(pattern, original_value, flags=re.IGNORECASE):
            return replacement

    # Priority 2: Generic suffix replacements for standard machinery types
    suffix_mapping = {
        r"(.*)(?:Aft)$": r"\1A",
        r"(.*)(?:Forward)$": r"\1F",
        r"(.*)(?:Fwd)$": r"\1F",
        r"(.*)(?:Port)$": r"\1P",
        r"(.*)(?:Starboard)$": r"\1S",
        r"(.*)(?:-P)$": r"\1P",
        r"(.*)(?:-S)$": r"\1S",
        r"(.*)(?:-Port)$": r"\1P",
        r"(.*)(?:-Stbd)$": r"\1S",
    }

    for pattern, replacement in suffix_mapping.items():
        if re.match(pattern, original_value, flags=re.IGNORECASE):
            return re.sub(pattern, replacement, original_value, flags=re.IGNORECASE).strip()

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


