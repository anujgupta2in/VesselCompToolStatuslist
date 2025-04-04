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

    # Lifeboat
    r"^LifeboatA1$": "Lifeboat A1",
    r"^LifeboatAft1$": "Lifeboat A1",

    # Liferaft Embarkation Ladder
    r"^Liferaft Embarkation LadderF1$": "Liferaft Embarkation Ladder F1",
    r"^Liferaft Embarkation LadderForward1$": "Liferaft Embarkation Ladder F1",
    r"^Liferaft Embarkation LadderP1$": "Liferaft Embarkation Ladder P1",
    r"^Liferaft Embarkation LadderPort1$": "Liferaft Embarkation Ladder P1",
    r"^Liferaft Embarkation LadderS1$": "Liferaft Embarkation Ladder S1",
    r"^Liferaft Embarkation LadderStarboard1$": "Liferaft Embarkation Ladder S1",

    # Liferaft
    r"^LiferaftP1$": "Liferaft P1",
    r"^LiferaftPort1$": "Liferaft P1",
    r"^LiferaftP2$": "Liferaft P2",
    r"^LiferaftPort2$": "Liferaft P2",
    r"^LiferaftS1$": "Liferaft S1",
    r"^LiferaftStarboard1$": "Liferaft S1",
    r"^LiferaftS2$": "Liferaft S2",
    r"^LiferaftStarboard2$": "Liferaft S2",

    # Mooring Winch
    r"^Mooring WinchA3$": "Mooring Winch A3",
    r"^Mooring WinchAft3$": "Mooring Winch A3",
    r"^Mooring WinchA4$": "Mooring Winch A4",
    r"^Mooring WinchAft4$": "Mooring Winch A4",
    r"^Mooring WinchF1$": "Mooring Winch F1",
    r"^Mooring WinchForward1$": "Mooring Winch F1",
    r"^Mooring WinchF2$": "Mooring Winch F2",
    r"^Mooring WinchForward2$": "Mooring Winch F2",

    # Pilot Ladder
    r"^Pilot LadderP1$": "Pilot Ladder P1",
    r"^Pilot LadderPort1$": "Pilot Ladder P1",
    r"^Pilot LadderS1$": "Pilot Ladder S1",
    r"^Pilot LadderStarboard1$": "Pilot Ladder S1",

    # Rescue Boat
    r"^Rescue BoatP1$": "Rescue Boat P1",
    r"^Rescue BoatPort1$": "Rescue Boat P1",

    r"^Combined Mooring Winch Hydraulic UnitF1$": "Combined Mooring Winch Hydraulic Unit F1",
    r"^Combined Mooring Winch Hydraulic UnitForward1$": "Combined Mooring Winch Hydraulic Unit F1",

    # Emergency Towing System
    r"^Emergency Towing SystemA1$": "Emergency Towing System A1",
    r"^Emergency Towing SystemAft1$": "Emergency Towing System A1",
    r"^Emergency Towing SystemF1$": "Emergency Towing System F1",
    r"^Emergency Towing SystemForward1$": "Emergency Towing System F1",

    # Liferaft 15
    r"^Liferaft 15P1$": "Liferaft 15P1",
    r"^Liferaft 15P2$": "Liferaft 15P2",
    r"^Liferaft 15Port1$": "Liferaft 15P1",
    r"^Liferaft 15Port2$": "Liferaft 15P2",

    # Liferaft 6PF
    r"^Liferaft 6PF-P1$": "Liferaft 6PF-P1",
    r"^Liferaft 6PFwd-Port1$": "Liferaft 6PF-P1",

    # Liferaft Embarkation Ladder F-*
    r"^Liferaft Embarkation LadderF-P1$": "Liferaft Embarkation Ladder F-P1",
    r"^Liferaft Embarkation LadderF-S1$": "Liferaft Embarkation Ladder F-S1",
    r"^Liferaft Embarkation LadderFwd-Port1$": "Liferaft Embarkation Ladder F-P1",
    r"^Liferaft Embarkation LadderFwd-Stbd1$": "Liferaft Embarkation Ladder F-S1",

    # Mooring Winch Hydraulic Unit
    r"^Mooring Winch Hydraulic UnitA1$": "Mooring Winch Hydraulic Unit A1",
    r"^Mooring Winch Hydraulic UnitAft1$": "Mooring Winch Hydraulic Unit A1",

    # Rescue Boat S
    r"^Rescue BoatS1$": "Rescue Boat S1",
    r"^Rescue BoatStarboard1$": "Rescue Boat S1",

    # SART
    r"^SARTP1$": "SART P1",
    r"^SARTPort1$": "SART P1",
    r"^SARTS1$": "SART S1",
    r"^SARTStarboard1$": "SART S1",

        # Liferaft 15PPort
    r"^Liferaft 15PPort1$": "Liferaft 15PP1",
    r"^Liferaft 15PPort2$": "Liferaft 15PP2", 

         # ICCP
    r"^ICCPA1$": "ICCP A1",
    r"^ICCPAft1$": "ICCP A1",
    r"^ICCPF1$": "ICCP F1",
    r"^ICCPForward1$": "ICCP F1",

    # Slewing Fuel Hose Crane
    r"^Slewing Fuel Hose CraneP1$": "Slewing Fuel Hose Crane P1",
    r"^Slewing Fuel Hose CranePort1$": "Slewing Fuel Hose Crane P1",
    r"^Slewing Fuel Hose CraneS1$": "Slewing Fuel Hose Crane S1",
    r"^Slewing Fuel Hose CraneStarboard1$": "Slewing Fuel Hose Crane S1",

     # Combined Windlass Mooring Winch F-*
    r"^Combined Windlass Mooring WinchF-P1$": "Combined Windlass Mooring Winch F-P1",
    r"^Combined Windlass Mooring WinchF-S1$": "Combined Windlass Mooring Winch F-S1",
    r"^Combined Windlass Mooring WinchFwd-Port1$": "Combined Windlass Mooring Winch F-P1",
    r"^Combined Windlass Mooring WinchFwd-Stbd1$": "Combined Windlass Mooring Winch F-S1",

    # Lifeboat Davit
    r"^Lifeboat DavitP1$": "Lifeboat Davit P1",
    r"^Lifeboat DavitPort1$": "Lifeboat Davit P1",

    # Lifeboat
    r"^LifeboatP1$": "Lifeboat P1",
    r"^LifeboatPort1$": "Lifeboat P1",

    # Liferaft Embarkation Ladder P2/S2
    r"^Liferaft Embarkation LadderP2$": "Liferaft Embarkation Ladder P2",
    r"^Liferaft Embarkation LadderPort2$": "Liferaft Embarkation Ladder P2",
    r"^Liferaft Embarkation LadderS2$": "Liferaft Embarkation Ladder S2",
    r"^Liferaft Embarkation LadderStarboard2$": "Liferaft Embarkation Ladder S2",

    # Liferaft/Rescue Boat Davit
    r"^Liferaft/Rescue Boat DavitS1$": "Liferaft/Rescue Boat Davit S1",
    r"^Liferaft/Rescue Boat DavitStarboard1$": "Liferaft/Rescue Boat Davit S1",

    # Mooring Winch Centre
    r"^Mooring WinchC1$": "Mooring Winch C1",
    r"^Mooring WinchCentre1$": "Mooring Winch C1",


       # Hatch Cover Aft to A mapping
    r"^Hatch CoverA1$": "Hatch Cover A1",
    r"^Hatch CoverA2$": "Hatch Cover A2",
    r"^Hatch CoverA3$": "Hatch Cover A3",
    r"^Hatch CoverA4$": "Hatch Cover A4",
    r"^Hatch CoverA5$": "Hatch Cover A5",
    r"^Hatch CoverA6$": "Hatch Cover A6",
    r"^Hatch CoverA7$": "Hatch Cover A7",
    r"^Hatch CoverAft1$": "Hatch Cover A1",
    r"^Hatch CoverAft2$": "Hatch Cover A2",
    r"^Hatch CoverAft3$": "Hatch Cover A3",
    r"^Hatch CoverAft4$": "Hatch Cover A4",
    r"^Hatch CoverAft5$": "Hatch Cover A5",
    r"^Hatch CoverAft6$": "Hatch Cover A6",
    r"^Hatch CoverAft7$": "Hatch Cover A7",

    # Hatch Cover Centre to C mapping
    r"^Hatch CoverC1$": "Hatch Cover C1",
    r"^Hatch CoverC2$": "Hatch Cover C2",
    r"^Hatch CoverCentre1$": "Hatch Cover C1",
    r"^Hatch CoverCentre2$": "Hatch Cover C2",

    # Hatch Cover Forward to F mapping
    r"^Hatch CoverF1$": "Hatch Cover F1",
    r"^Hatch CoverF2$": "Hatch Cover F2",
    r"^Hatch CoverF3$": "Hatch Cover F3",
    r"^Hatch CoverF4$": "Hatch Cover F4",
    r"^Hatch CoverF5$": "Hatch Cover F5",
    r"^Hatch CoverF6$": "Hatch Cover F6",
    r"^Hatch CoverF7$": "Hatch Cover F7",
    r"^Hatch CoverForward1$": "Hatch Cover F1",
    r"^Hatch CoverForward2$": "Hatch Cover F2",
    r"^Hatch CoverForward3$": "Hatch Cover F3",
    r"^Hatch CoverForward4$": "Hatch Cover F4",
    r"^Hatch CoverForward5$": "Hatch Cover F5",
    r"^Hatch CoverForward6$": "Hatch Cover F6",
    r"^Hatch CoverForward7$": "Hatch Cover F7",

  # Mooring Winch Centre to C mapping (corrected)
    r"^Mooring WinchC2$": "Mooring Winch C2",
    r"^Mooring WinchCentre2$": "Mooring Winch C2",

    # Mooring Winch P variants
    r"^Mooring WinchP1$": "Mooring Winch P1",
    r"^Mooring WinchP2$": "Mooring Winch P2",
    r"^Mooring WinchP3$": "Mooring Winch P3",
    r"^Mooring WinchPort1$": "Mooring Winch P1",
    r"^Mooring WinchPort2$": "Mooring Winch P2",
    r"^Mooring WinchPort3$": "Mooring Winch P3",

    # Mooring Winch S variants
    r"^Mooring WinchS1$": "Mooring Winch S1",
    r"^Mooring WinchS2$": "Mooring Winch S2",
    r"^Mooring WinchStarboard1$": "Mooring Winch S1",
    r"^Mooring WinchStarboard2$": "Mooring Winch S2",


    # Lifeboat/Rescue Boat
    r"^Lifeboat/Rescue BoatS1$": "Lifeboat/Rescue Boat S1",
    r"^Lifeboat/Rescue BoatStarboard1$": "Lifeboat/Rescue Boat S1",

    # Liferaft F1 / Forward1
    r"^LiferaftF1$": "Liferaft F1",
    r"^LiferaftForward1$": "Liferaft F1",

    # Muster Station
    r"^Muster StationP1$": "Muster Station P1",
    r"^Muster StationPort1$": "Muster Station P1",
    r"^Muster StationS1$": "Muster Station S1",
    r"^Muster StationStarboard1$": "Muster Station S1",

        # Pilot Combination Ladder P2
    r"^Pilot Combination LadderP2$": "Pilot Combination Ladder P2",
    r"^Pilot Combination LadderPort2$": "Pilot Combination Ladder P2",

  
    # Liferaft Forward Port/Starboard
    r"^LiferaftFP$": "Liferaft FP",
    r"^LiferaftFS$": "Liferaft FS",
    r"^LiferaftFwd-P$": "Liferaft FP",
    r"^LiferaftFwdS$": "Liferaft FS",

    # Lifeboat Davit
    r"^Lifeboat DavitS1$": "Lifeboat Davit S1",
    r"^Lifeboat DavitStarboard1$": "Lifeboat Davit S1",

    # Lifeboat/Rescue Boat
    r"^Lifeboat/Rescue BoatP1$": "Lifeboat/Rescue Boat P1",
    r"^Lifeboat/Rescue BoatPort1$": "Lifeboat/Rescue Boat P1",

    # Lifeboat
    r"^LifeboatS1$": "Lifeboat S1",
    r"^LifeboatStarboard1$": "Lifeboat S1",

    # Liferaft 16 Person
    r"^Liferaft 16 PersonP1$": "Liferaft 16 Person P1",
    r"^Liferaft 16 PersonP2$": "Liferaft 16 Person P2",
    r"^Liferaft 16 PersonPort1$": "Liferaft 16 Person P1",
    r"^Liferaft 16 PersonPort2$": "Liferaft 16 Person P2",
    r"^Liferaft 16 PersonS1$": "Liferaft 16 Person S1",
    r"^Liferaft 16 PersonS2$": "Liferaft 16 Person S2",
    r"^Liferaft 16 PersonStarboard1$": "Liferaft 16 Person S1",
    r"^Liferaft 16 PersonStarboard2$": "Liferaft 16 Person S2",

    # Liferaft 6 Person
    r"^Liferaft 6 PersonF-P1$": "Liferaft 6 Person F-P1",
    r"^Liferaft 6 PersonFwd-Port1$": "Liferaft 6 Person F-P1",

    # Liferaft/Rescue Boat Davit
    r"^Liferaft/Rescue Boat DavitP1$": "Liferaft/Rescue Boat Davit P1",
    r"^Liferaft/Rescue Boat DavitPort1$": "Liferaft/Rescue Boat Davit P1",

    # Mooring Winch M
    r"^Mooring WinchM1$": "Mooring Winch M1",
    r"^Mooring WinchM2$": "Mooring Winch M2",
    r"^Mooring WinchM3$": "Mooring Winch M3",
    r"^Mooring WinchM4$": "Mooring Winch M4",
    r"^Mooring WinchM5$": "Mooring Winch M5",
    r"^Mooring WinchM6$": "Mooring Winch M6",
    r"^Mooring WinchMiddle1$": "Mooring Winch M1",
    r"^Mooring WinchMiddle2$": "Mooring Winch M2",
    r"^Mooring WinchMiddle3$": "Mooring Winch M3",
    r"^Mooring WinchMiddle4$": "Mooring Winch M4",
    r"^Mooring WinchMiddle5$": "Mooring Winch M5",
    r"^Mooring WinchMiddle6$": "Mooring Winch M6",

        # Liferaft/Rescue Boat Davit S2
    r"^Liferaft/Rescue Boat DavitS2$": "Liferaft/Rescue Boat Davit S2",
    r"^Liferaft/Rescue Boat DavitStarboard2$": "Liferaft/Rescue Boat Davit S2",

        # Lifeboat/Rescue Boat Davit S1
    r"^Lifeboat/Rescue Boat DavitS1$": "Lifeboat/Rescue Boat Davit S1",
    r"^Lifeboat/Rescue Boat DavitStarboard1$": "Lifeboat/Rescue Boat Davit S1",

    # Liferaft Embarkation Ladder P3/S3
    r"^Liferaft Embarkation LadderP3$": "Liferaft Embarkation Ladder P3",
    r"^Liferaft Embarkation LadderPort3$": "Liferaft Embarkation Ladder P3",
    r"^Liferaft Embarkation LadderS3$": "Liferaft Embarkation Ladder S3",
    r"^Liferaft Embarkation LadderStarboard3$": "Liferaft Embarkation Ladder S3",

        # Liferaft 6 Person F1
    r"^Liferaft 6 PersonF1$": "Liferaft 6 Person F1",
    r"^Liferaft 6 PersonForward1$": "Liferaft 6 Person F1",

    # Mooring Winch Aft combinations
    r"^Mooring WinchA-P1$": "Mooring Winch A-P1",
    r"^Mooring WinchA-P2$": "Mooring Winch A-P2",
    r"^Mooring WinchA-S1$": "Mooring Winch A-S1",
    r"^Mooring WinchA-S2$": "Mooring Winch A-S2",
    r"^Mooring WinchAft-Port1$": "Mooring Winch A-P1",
    r"^Mooring WinchAft-Port2$": "Mooring Winch A-P2",
    r"^Mooring WinchAft-Stbd1$": "Mooring Winch A-S1",
    r"^Mooring WinchAft-Stbd2$": "Mooring Winch A-S2",

    # Mooring Winch Forward combinations
    r"^Mooring WinchF-P1$": "Mooring Winch F-P1",
    r"^Mooring WinchF-S1$": "Mooring Winch F-S1",
    r"^Mooring WinchFwd-Port1$": "Mooring Winch F-P1",
    r"^Mooring WinchFwd-Stbd1$": "Mooring Winch F-S1",


        # Liferaft FP / FS
    r"^LiferaftFP$": "Liferaft FP",
    r"^LiferaftFS$": "Liferaft FS",
    r"^LiferaftFwd-P$": "Liferaft FP",
    r"^LiferaftFwdS$": "Liferaft FS",

    # Combined Mooring Winch Hydraulic Unit
    r"^Combined Mooring Winch Hydraulic UnitA1$": "Combined Mooring Winch Hydraulic Unit A1",
    r"^Combined Mooring Winch Hydraulic UnitAft1$": "Combined Mooring Winch Hydraulic Unit A1",

    # Emergency Towing System F2
    r"^Emergency Towing SystemF2$": "Emergency Towing System F2",
    r"^Emergency Towing SystemForward2$": "Emergency Towing System F2",

    # Liferaft 20 Person
    r"^Liferaft 20 PersonP1$": "Liferaft 20 Person P1",
    r"^Liferaft 20 PersonP2$": "Liferaft 20 Person P2",
    r"^Liferaft 20 PersonPort1$": "Liferaft 20 Person P1",
    r"^Liferaft 20 PersonPort2$": "Liferaft 20 Person P2",
    r"^Liferaft 20 PersonS1$": "Liferaft 20 Person S1",
    r"^Liferaft 20 PersonS2$": "Liferaft 20 Person S2",
    r"^Liferaft 20 PersonStarboard1$": "Liferaft 20 Person S1",
    r"^Liferaft 20 PersonStarboard2$": "Liferaft 20 Person S2",

    # Mooring Winch Hydraulic Unit Forward
    r"^Mooring Winch Hydraulic UnitF1$": "Mooring Winch Hydraulic Unit F1",
    r"^Mooring Winch Hydraulic UnitForward1$": "Mooring Winch Hydraulic Unit F1",

    # Provision Crane Starboard
    r"^Provision Crane StbdS1$": "Provision Crane S1",
    r"^Provision Crane StbdStarboard1$": "Provision Crane S1",

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


