import streamlit as st

st.title("🎈 AZURE NSG Rules - JSON to nice Excel coverter")
st.write(
    "Convert Azure NSG JSON export to Excel"
)

import io
import json
import re
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- Region formatter ---
def format_location(loc: str) -> str:
    """Convert Azure location code into a human-friendly region name."""
    if not loc:
        return ""
    loc_lower = loc.lower()
    region_map = {
        "australiaeast": "Australia East",
        "australiasoutheast": "Australia Southeast",
        "australiacentral": "Australia Central",
        "australiacentral2": "Australia Central 2",
        "southeastasia": "Southeast Asia",
        "eastasia": "East Asia",
        "japaneast": "Japan East",
        "japanwest": "Japan West",
        "koreacentral": "Korea Central",
        "koreasouth": "Korea South",
        "southindia": "South India",
        "centralindia": "Central India",
        "westindia": "West India",
        # ... (other regions)
    }

    if loc_lower in region_map:
        return region_map[loc_lower]
    loc_cleaned = re.sub(r'([a-z])([A-Z0-9])', r'\1 \2', loc_lower.title())
    loc_cleaned = loc_cleaned.replace("Azure ", "").replace("-", " ").title()
    return loc_cleaned

# --- Streamlit File Upload ---
st.title("Convert Azure NSG JSON to Excel")
uploaded_file = st.file_uploader("Choose a JSON file", type="json")

if uploaded_file is not None:
    # --- Parse JSON ---
    data = json.load(uploaded_file)
    nsg_name = data.get("name", "Unknown NSG")
    id_path = data.get("id", "")
    subscription_id = id_path.split("/")[2] if id_path else ""
    resource_group = id_path.split("/resourceGroups/")[1].split("/")[0] if "/resourceGroups/" in id_path else ""
    metadata = {
        "Resource group": resource_group,
        "Location": format_location(data.get("location", "")),
        "Subscription ID": subscription_id,
    }

    # --- RULES ---
    rules = data["properties"].get("securityRules", []) + data["properties"].get("defaultSecurityRules", [])
    
    def replace_any(v):
        if isinstance(v, str):
            return "Any" if v.strip() == "*" else v
        if isinstance(v, list):
            return ", ".join(replace_any(x) for x in v if x)
        return v

    records = []
    for r in rules:
        p = r.get("properties", {})
        ports = p.get("destinationPortRanges", []) or [p.get("destinationPortRange", "Any")]
        sources = p.get("sourceAddressPrefixes", []) or [p.get("sourceAddressPrefix", "Any")]
        dests = p.get("destinationAddressPrefixes", []) or [p.get("destinationAddressPrefix", "Any")]
        records.append({
            "Priority": int(p.get("priority", 0)),
            "Direction": p.get("direction", ""),
            "RuleName": r.get("name", ""),
            "Port": replace_any(ports),
            "Protocol": replace_any(p.get("protocol", "")),
            "Source": replace_any(sources),
            "Destination": replace_any(dests),
            "Access": p.get("access", ""),
            "Description": p.get("description", "")
        })

    df_rules = pd.DataFrame(records)
    if not df_rules.empty:
        df_rules["Direction"] = pd.Categorical(df_rules["Direction"], categories=["Inbound", "Outbound"], ordered=True)
        df_rules = df_rules.sort_values(["Direction", "Priority"])

    # --- Excel File Creation ---
    wb = Workbook()
    ws = wb.active
    ws.title = "NSG_RULES"
    bold = Font(bold=True)
    title_font = Font(bold=True, size=14)
    hdr_fill = PatternFill("solid", "D9E1F2")
    title_fill = PatternFill("solid", "BDD7EE")
    align_center = Alignment(horizontal="center")
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    # Add NSG name and metadata to Excel
    ws.append([nsg_name])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = align_center
    ws.append([""])
    for k, v in metadata.items():
        ws.append([k, v])
        ws[f"A{ws.max_row}"].font = bold
    ws.append([""])
    
    # Add Rules Table to Excel
    ws.append(["ROUTES"])
    ws.append(df_rules.columns.tolist())
    for i in range(1, 9):
        cell = ws[f"{get_column_letter(i)}{ws.max_row}"]
        cell.font = bold
        cell.fill = hdr_fill
    for r in df_rules.itertuples(index=False):
        ws.append(list(r))

    # Add border and column width adjustments
    for row in ws.iter_rows():
        for c in row:
            if c.value:
                c.border = border
    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(len(str(c.value)) for c in col if c.value) + 3

    # Save Excel file to BytesIO object
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)

    # --- Display Tables on Streamlit UI ---
    st.subheader("NSG Rules Table")
    st.dataframe(df_rules)

    # --- Streamlit Save As File ---
    st.download_button(
        label="Download Excel",
        data=excel_file,
        file_name=f"{nsg_name}_NSG_Rules.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
