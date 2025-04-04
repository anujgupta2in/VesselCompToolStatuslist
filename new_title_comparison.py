import pandas as pd
import streamlit as st
from io import BytesIO
import re
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

def extract_date_from_filename(filename):
    """Extract and format date from filename."""
    base_name = os.path.splitext(os.path.basename(filename))[0]
    date_part = base_name.split()[-1]
    try:
        if len(date_part) >= 8:  # Format like 25032025
            return f"{date_part[0:2]}-{date_part[2:4]}-{date_part[4:8]}"
        else:
            return date_part
    except:
        return date_part

def get_vessel_name(df):
    """Extract vessel name from DataFrame."""
    if "Vessel" in df.columns:
        vessel_values = df["Vessel"].dropna()
        if not vessel_values.empty:
            return vessel_values.iloc[0]
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



def prepare_excel_report(df, file1_name, file2_name, vessel1_name, vessel2_name):
    """Create a formatted Excel report based on the comparison results."""
    
    # Create a new workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Job Title Comparison"
    
    if not df.empty:
        # Write header for main comparison sheet
        headers = df.columns.tolist()
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # Write data for main comparison sheet
        for row_idx, row_data in enumerate(df.values, 2):
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Define styles
    fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # Light yellow
    fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Light red
    fill_light_blue = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")  # Light blue
    bold_font = Font(bold=True)
    red_font = Font(color="9C0006")  # Dark red
    
    # Apply styling to main comparison sheet
    if not df.empty:
        # Format headers
        for col in range(1, len(df.columns) + 1):
            ws.cell(row=1, column=col).font = bold_font
        
        # Format data rows
        for row in range(2, len(df) + 2):
            # Get the "Has Differences" value - safely handle missing columns
            has_diff_col = None
            for i, col_name in enumerate(df.columns):
                if col_name == 'Has Differences':
                    has_diff_col = i
                    break
            
            # Only proceed if we found the column
            if has_diff_col is not None:
                has_diff = ws.cell(row=row, column=has_diff_col + 1).value
                
                if has_diff == 'Yes':
                    # Highlight machinery name
                    ws.cell(row=row, column=1).font = bold_font
                    
                    # Find title columns
                    for col in range(1, len(df.columns) + 1):
                        header = ws.cell(row=1, column=col).value
                        if 'Titles only in' in header:
                            cell = ws.cell(row=row, column=col)
                            if cell.value != '-':
                                cell.fill = fill_yellow
                    
                    # Highlight "Has Differences" column
                    ws.cell(row=row, column=has_diff_col + 1).font = red_font
                    ws.cell(row=row, column=has_diff_col + 1).fill = fill_red
        
        # Adjust column widths and set text wrapping
        for col in range(1, len(df.columns) + 1):
            col_letter = chr(64 + col)  # Convert column number to letter (A, B, C, ...)
            ws.column_dimensions[col_letter].width = 30
            
            # Set text wrapping for all cells
            for row in range(2, len(df) + 2):
                cell = ws.cell(row=row, column=col)
                cell.alignment = Alignment(wrap_text=True, vertical='top')
    
    # Add a separate sheet for just the list of machinery with differences
    machinery_diff_sheet = wb.create_sheet(title="Machinery Differences")
    
    # Check if we have any machinery with differences
    diff_machinery = []
    if not df.empty:
        diff_machinery = df[df['Has Differences'] == 'Yes']['Machinery'].tolist()
    
    # Add a header row for the machinery differences sheet
    machinery_diff_sheet.cell(row=1, column=1, value="Machinery with Different Job Titles")
    machinery_diff_sheet.cell(row=1, column=2, value=f"Comparison: {vessel1_name} vs {vessel2_name}")
    machinery_diff_sheet.cell(row=1, column=1).font = bold_font
    machinery_diff_sheet.cell(row=1, column=2).font = bold_font
    
    # Add a subheader row
    machinery_diff_sheet.cell(row=3, column=1, value="No.")
    machinery_diff_sheet.cell(row=3, column=2, value="Machinery")
    machinery_diff_sheet.cell(row=3, column=1).font = bold_font
    machinery_diff_sheet.cell(row=3, column=2).font = bold_font
    
    # Add the machinery names to the sheet
    for idx, machinery in enumerate(sorted(diff_machinery), 1):
        machinery_diff_sheet.cell(row=idx+3, column=1, value=idx)
        machinery_diff_sheet.cell(row=idx+3, column=2, value=machinery)
        # Apply alternating row coloring for better readability
        if idx % 2 == 0:
            machinery_diff_sheet.cell(row=idx+3, column=1).fill = fill_light_blue
            machinery_diff_sheet.cell(row=idx+3, column=2).fill = fill_light_blue
    
    # Make the second column wider
    machinery_diff_sheet.column_dimensions['B'].width = 50
    
    # If no machinery with differences found, add a note
    if not diff_machinery:
        machinery_diff_sheet.cell(row=4, column=1, value="No machinery with different job titles found")
        machinery_diff_sheet.cell(row=4, column=1).font = Font(italic=True)
    
    # Make sure at least one sheet is visible
    if len(wb.sheetnames) > 0:
        # Set all sheets to visible state
        for sheet_name in wb.sheetnames:
            wb[sheet_name].sheet_state = 'visible'
    else:
        # Create a blank sheet if none exists
        ws = wb.create_sheet("No Differences")
        ws.append(["No job title differences found between the two files"])
    
    # Save the styled workbook with error handling
    try:
        output_final = BytesIO()
        wb.save(output_final)
        output_final.seek(0)
        return output_final.getvalue()
    except Exception as e:
        print(f"Error saving Excel file: {str(e)}")
        # Return a simple Excel file with error message
        wb_error = Workbook()
        ws_error = wb_error.active
        ws_error.title = "Error"
        ws_error.append(["Error generating report", str(e)])
        output_error = BytesIO()
        wb_error.save(output_error)
        output_error.seek(0)
        return output_error.getvalue()

def compare_titles(file1_content, file2_content, file1_name, file2_name):
    """Compare job titles between two CSV files for each machinery."""
    try:
        # Read CSV files
        df1 = pd.read_csv(BytesIO(file1_content))
        df2 = pd.read_csv(BytesIO(file2_content))
        
        # Print column names for debugging
        print("First file columns:", df1.columns.tolist())
        print("Second file columns:", df2.columns.tolist())
        
        # Extract dates and vessel names
        date1_fmt = extract_date_from_filename(file1_name)
        date2_fmt = extract_date_from_filename(file2_name)
        vessel1 = get_vessel_name(df1)
        vessel2 = get_vessel_name(df2)
        
        # Determine columns based on actual file structure
        first_machinery_col = None
        first_title_col = None
        second_machinery_col = None
        second_title_col = None
        
        # Check first file
        if 'Machinery Location' in df1.columns:
            first_machinery_col = 'Machinery Location'
        elif 'Machinery' in df1.columns:
            first_machinery_col = 'Machinery'
        
        if 'Title' in df1.columns:
            first_title_col = 'Title'
        elif 'Job Title' in df1.columns:
            first_title_col = 'Job Title'
        # Check for Job Title.1
        elif 'Job Title.1' in df1.columns:
            first_title_col = 'Job Title.1'
        
        # Check second file  
        if 'Machinery Location' in df2.columns:
            second_machinery_col = 'Machinery Location'
        elif 'Machinery' in df2.columns:
            second_machinery_col = 'Machinery'
        
        # Handle job title columns in the second file
        if 'Job Title' in df2.columns:
            # Prioritize Job Title column
            second_title_col = 'Job Title'
        elif 'Title' in df2.columns:
            second_title_col = 'Title'
        elif 'Job Title.1' in df2.columns:
            second_title_col = 'Job Title.1'
        
        if first_machinery_col is None:
            raise ValueError("Machinery column not found in first file. Available columns: " + 
                            str(df1.columns.tolist()))
        
        if first_title_col is None:
            raise ValueError("Title/Job Title column not found in first file. Available columns: " + 
                            str(df1.columns.tolist()))
        
        if second_machinery_col is None:
            raise ValueError("Machinery column not found in second file. Available columns: " + 
                            str(df2.columns.tolist()))
        
        if second_title_col is None:
            raise ValueError("Title/Job Title column not found in second file. Available columns: " + 
                            str(df2.columns.tolist()))
        
        print(f"Using columns: {first_machinery_col}, {first_title_col} from first file")
        print(f"Using columns: {second_machinery_col}, {second_title_col} from second file")
        
        # Print the first few rows of data for debugging
        print("\nFirst file sample data:")
        for idx, row in df1.head(3).iterrows():
            print(f"  Row {idx}: {first_machinery_col}={row[first_machinery_col]}, {first_title_col}={row[first_title_col]}")
        
        print("\nSecond file sample data:")
        for idx, row in df2.head(3).iterrows():
            print(f"  Row {idx}: {second_machinery_col}={row[second_machinery_col]}, {second_title_col}={row[second_title_col]}")
        
        # Standardize machinery names
        df1[first_machinery_col] = df1[first_machinery_col].apply(lambda x: rename_machinery(str(x)) if pd.notna(x) else x)
        df2[second_machinery_col] = df2[second_machinery_col].apply(lambda x: rename_machinery(str(x)) if pd.notna(x) else x)
        
        # Format column names for display
        col1 = f"{vessel1} ({date1_fmt})"
        col2 = f"{vessel2} ({date2_fmt})"
        
        # Prepare dataframes for title comparison
        titles_df1 = df1[[first_machinery_col, first_title_col]].copy()
        titles_df1.rename(columns={first_machinery_col: 'Machinery', first_title_col: 'Job Title'}, inplace=True)
        titles_df1.drop_duplicates(inplace=True)
        
        titles_df2 = df2[[second_machinery_col, second_title_col]].copy()
        titles_df2.rename(columns={second_machinery_col: 'Machinery', second_title_col: 'Job Title'}, inplace=True)
        titles_df2.drop_duplicates(inplace=True)
        
        # Filter out any rows with missing machinery names
        titles_df1 = titles_df1[titles_df1['Machinery'].notna()]
        titles_df2 = titles_df2[titles_df2['Machinery'].notna()]
        
        # Convert all to strings for comparison
        titles_df1['Machinery'] = titles_df1['Machinery'].astype(str)
        titles_df1['Job Title'] = titles_df1['Job Title'].astype(str)
        titles_df2['Machinery'] = titles_df2['Machinery'].astype(str)
        titles_df2['Job Title'] = titles_df2['Job Title'].astype(str)
        
        # Create a dictionary to store title comparison results
        title_comparison_results = []
        
        # Get unique machinery names from both DataFrames
        all_machinery = pd.concat([
            titles_df1['Machinery'], 
            titles_df2['Machinery']
        ]).drop_duplicates().tolist()
        
        # Compare titles for each machinery
        for machinery in all_machinery:
            if machinery == 'TOTAL':
                continue
                
            # Get titles from both DataFrames
            titles1 = titles_df1[titles_df1['Machinery'] == machinery]['Job Title'].tolist()
            titles2 = titles_df2[titles_df2['Machinery'] == machinery]['Job Title'].tolist()
            
            # Remove "nan" values from comparison
            titles1 = [t for t in titles1 if t != "nan"]
            titles2 = [t for t in titles2 if t != "nan"]
            
            # Find unique titles in each dataframe
            only_in_df1 = list(set(titles1) - set(titles2))
            only_in_df2 = list(set(titles2) - set(titles1))
            common_titles = list(set(titles1) & set(titles2))
            
            # Ensure unique column names by adding file names if vessel names are identical
            first_file_name = os.path.splitext(os.path.basename(file1_name))[0]
            second_file_name = os.path.splitext(os.path.basename(file2_name))[0]
            
            # Create column names for title differences - ensuring uniqueness
            if vessel1 == vessel2:
                first_title_col = f'Titles only in {vessel1} (File 1)'
                second_title_col = f'Titles only in {vessel2} (File 2)'
            else:
                first_title_col = f'Titles only in {vessel1}'
                second_title_col = f'Titles only in {vessel2}'
            
            # Create result dictionary with consistent columns
            # Fix the "Has Differences" flag logic - consider a difference if any titles exist in only one set
            # This handles the case where a machinery has titles in only one file (titles1 or titles2 empty)
            result_dict = {
                'Machinery': machinery,
                'Has Differences': 'Yes' if only_in_df1 or only_in_df2 else 'No',
                'Common Titles': ', '.join(sorted(common_titles)) if common_titles else '-',
                first_title_col: ', '.join(sorted(only_in_df1)) if only_in_df1 else '-',
                second_title_col: ', '.join(sorted(only_in_df2)) if only_in_df2 else '-'
            }
            
            # Include all machinery items with titles in at least one file
            if titles1 or titles2:
                title_comparison_results.append(result_dict)
        
        # Create DataFrame from results
        title_comparison_df = pd.DataFrame(title_comparison_results)
        
        # Create column names for title differences - ensuring uniqueness
        if vessel1 == vessel2:
            first_title_col = f'Titles only in {vessel1} (File 1)'
            second_title_col = f'Titles only in {vessel2} (File 2)'
        else:
            first_title_col = f'Titles only in {vessel1}'
            second_title_col = f'Titles only in {vessel2}'
            
        # If we have no comparison results, create an empty dataframe with the expected columns
        if title_comparison_df.empty:
            title_comparison_df = pd.DataFrame(columns=[
                'Machinery',
                'Has Differences',
                'Common Titles',
                first_title_col, 
                second_title_col
            ])
        else:
            # Ensure columns are in the correct order if all exist in the dataframe
            column_order = [
                'Machinery',
                'Has Differences',
                'Common Titles'
            ]
            
            # Add title columns that exist in the dataframe
            for col in title_comparison_df.columns:
                if 'Titles only in' in col and col not in column_order:
                    column_order.append(col)
            
            # Reorder columns that exist
            title_comparison_df = title_comparison_df[column_order]
            
            # Sort by machinery name
            title_comparison_df.sort_values('Machinery', inplace=True)
        
        # Prepare list of machinery with differences
        machinery_with_diff = title_comparison_df[title_comparison_df['Has Differences'] == 'Yes']['Machinery'].tolist()
        
        # Create Excel file
        excel_data = prepare_excel_report(title_comparison_df, file1_name, file2_name, vessel1, vessel2)
        
        return title_comparison_df, machinery_with_diff, excel_data
    except Exception as e:
        # Log the error for debugging
        print(f"Error in compare_titles: {str(e)}")
        # Return empty results rather than raising an error
        empty_df = pd.DataFrame(columns=[
            'Machinery',
            'Has Differences',
            'Common Titles',
            'Titles only in File 1',
            'Titles only in File 2'
        ])
        return empty_df, [], BytesIO().getvalue()

def render_title_comparison_app():
    """Render the Streamlit app for job title comparison."""
    # Page config is now set in main.py

    st.title("🚢 Job Title Comparison Tool")

    st.markdown("""
    This tool compares job titles for machinery between two CSV files.

    **Instructions:**
    1. Upload two CSV files containing machinery data
    2. The files should have machinery and job title columns:
       - Acceptable machinery column names include: 'Machinery', 'Machinery Location'
       - Acceptable job title column names include: 'Job Title', 'Title'
       - The tool will automatically detect and use the correct columns
    3. The tool will analyze the data and show machinery with different job titles
    4. Download the highlighted Excel report for offline use
    """)

    # File upload section
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("First CSV File (System Management)")
        file1 = st.file_uploader("Upload System Management CSV", type=['csv'], key="file1")
        if file1:
            df1 = pd.read_csv(file1)
            st.write("Preview of first file:")
            st.dataframe(df1.head(), use_container_width=True)
            st.info(f"Total rows: {len(df1)}")
            
            # Check for machinery columns
            has_machinery_col = False
            for col_name in ['Machinery', 'Machinery Location']:
                if col_name in df1.columns:
                    has_machinery_col = True
                    break
                    
            # Check for title columns
            has_title_col = False
            for col_name in ['Title', 'Job Title']:
                if col_name in df1.columns:
                    has_title_col = True
                    break
            
            if has_machinery_col and has_title_col:
                st.success("✅ Required column types found!")
            else:
                missing = []
                if not has_machinery_col:
                    missing.append("Machinery column ('Machinery' or 'Machinery Location')")
                if not has_title_col:
                    missing.append("Title column ('Job Title' or 'Title')")
                st.error(f"❌ Missing required columns: {', '.join(missing)}")
                st.write("Available columns:", df1.columns.tolist())

    with col2:
        st.subheader("Second CSV File (PMS Jobs)")
        file2 = st.file_uploader("Upload PMS Jobs CSV", type=['csv'], key="file2")
        if file2:
            df2 = pd.read_csv(file2)
            st.write("Preview of second file:")
            st.dataframe(df2.head(), use_container_width=True)
            st.info(f"Total rows: {len(df2)}")
            
            # Check for machinery columns
            has_machinery_col = False
            for col_name in ['Machinery', 'Machinery Location']:
                if col_name in df2.columns:
                    has_machinery_col = True
                    break
                    
            # Check for title columns
            has_title_col = False
            for col_name in ['Title', 'Job Title']:
                if col_name in df2.columns:
                    has_title_col = True
                    break
            
            if has_machinery_col and has_title_col:
                st.success("✅ Required column types found!")
            else:
                missing = []
                if not has_machinery_col:
                    missing.append("Machinery column ('Machinery' or 'Machinery Location')")
                if not has_title_col:
                    missing.append("Title column ('Job Title' or 'Title')")
                st.error(f"❌ Missing required columns: {', '.join(missing)}")
                st.write("Available columns:", df2.columns.tolist())

    # Add a separator
    st.markdown("---")

    # Automatically process files when both are uploaded
    if file1 and file2:
        st.subheader("Comparison Report")
        
        # Automatically process when both files are uploaded
        with st.spinner("Analyzing files and generating report..."):
            st.info("Comparing job titles across files. This might take a few moments...")
            
            try:
                # Compare titles with robust error handling
                title_diff_df, machinery_diff_list, excel_data = compare_titles(
                    file1.getvalue(), file2.getvalue(), file1.name, file2.name
                )
                
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
                        st.subheader("🔎 Examples of Job Title Differences")
                        
                        # Sample up to 5 machinery items to show detailed differences
                        sample_count = min(5, len(diff_only_df))
                        if sample_count > 0:
                            st.write("Below are examples of machinery with different job titles:")
                            
                            sample_machines = diff_only_df['Machinery'].head(sample_count).tolist()
                            
                            for idx, machinery in enumerate(sample_machines):
                                row = diff_only_df[diff_only_df['Machinery'] == machinery].iloc[0]
                                
                                st.write(f"**{idx+1}. {machinery}**")
                                
                                # Use expander for better organization of content
                                with st.expander(f"View all title details for {machinery}", expanded=True):
                                    # Get the title columns
                                    title_cols = [col for col in diff_only_df.columns if col.startswith('Titles only in')]
                                    
                                    # Display common titles first (if any)
                                    st.write("**Common Titles:**")
                                    if 'Common Titles' in row and row['Common Titles'] != '-':
                                        st.markdown(
                                            f"<div style='background-color: #E8F5E9; padding: 10px; border-radius: 5px;'>{row['Common Titles']}</div>", 
                                            unsafe_allow_html=True
                                        )
                                    else:
                                        st.write("*None*")
                                    
                                    st.markdown("---")
                                    
                                    # Display titles from both files in separate sections
                                    cols = st.columns(2)
                                    
                                    # First file titles
                                    with cols[0]:
                                        if len(title_cols) > 0:
                                            first_col = title_cols[0]
                                            st.write(f"**{first_col}:**")
                                            if row[first_col] != '-':
                                                st.markdown(
                                                    f"<div style='background-color: #FFF3E0; padding: 10px; border-radius: 5px;'>{row[first_col]}</div>", 
                                                    unsafe_allow_html=True
                                                )
                                            else:
                                                st.write("*None*")
                                    
                                    # Second file titles
                                    with cols[1]:
                                        if len(title_cols) > 1:
                                            second_col = title_cols[1]
                                            st.write(f"**{second_col}:**")
                                            if row[second_col] != '-':
                                                st.markdown(
                                                    f"<div style='background-color: #E3F2FD; padding: 10px; border-radius: 5px;'>{row[second_col]}</div>", 
                                                    unsafe_allow_html=True
                                                )
                                            else:
                                                st.write("*None*")
                                
                                st.write("---")
                    else:
                        st.success("No job title differences found for any machinery!")
                else:
                    st.info("No job title comparison data generated. Please check if both files have matching machinery.")
                
                # Download section
                if isinstance(excel_data, bytes) and len(excel_data) > 0:
                    st.subheader("📥 Download Report")
                    st.write("Download the detailed Excel report with highlighted job title differences:")
                    st.download_button(
                        label="Download Job Title Comparison Report",
                        data=excel_data,
                        file_name="Job_Title_Comparison.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Click to download the job title comparison report in Excel format"
                    )
                
            except Exception as e:
                st.error(f"Error comparing job titles: {str(e)}")
                st.error("Please ensure both files have the required columns and proper format:")
                st.info("Acceptable column names: 'Machinery'/'Machinery Location' for machinery data")
                st.info("Acceptable column names: 'Job Title'/'Title' for job title data")
                st.info("The tool will automatically detect and use the appropriate columns if available")
    else:
        st.info("👆 Please upload both CSV files to generate the title comparison report")

    # Add footer
    st.markdown("---")
    st.markdown("Built with ❤️ using Streamlit")

if __name__ == "__main__":
    render_title_comparison_app()
