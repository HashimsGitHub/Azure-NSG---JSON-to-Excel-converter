
### 🛡️ NSG Rules — JSON to Excel Converter
**[`Azure_NSG_Rules_JSON_to_Excel_coverter.py`](./Azure_NSG_Rules_JSON_to_Excel_coverter.py)** · **[`Azure_NSG_Rules_JSON_to_Excel_coverter.ipynb`](./Azure_NSG_Rules_JSON_to_Excel_coverter.ipynb)**

🌐 **Live app:** [azure-nsg-json2excel.streamlit.app](http://azure-nsg-json2excel.streamlit.app/)

Converts an Azure Network Security Group (NSG) JSON export into a clean, formatted Excel spreadsheet. Inbound rules are listed first, followed by Outbound, both sorted by Priority. Includes NSG metadata (resource group, location, subscription ID) and human-readable Azure region names.

**Features:**
- GUI file picker — no command-line arguments needed
- Supports both custom and default security rules
- Expands multi-value port/address arrays into readable strings
- Replaces `*` wildcards with `Any` for clarity
- Full Azure region name mapping (60+ regions including Australia, US, Europe, Middle East, China)
- Styled Excel output with colour-coded headers

**Dependencies:**
```bash
pip install pandas openpyxl
```

**Usage:** Run the script and use the file picker to select your NSG JSON export (downloaded from the Azure Portal), then choose a save location for the `.xlsx` output.

---
