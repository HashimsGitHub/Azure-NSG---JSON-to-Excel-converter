<div align="center">

<img src="https://upload.wikimedia.org/wikipedia/commons/f/fa/Microsoft_Azure.svg" width="80" alt="Azure Logo"/>

# Azure NSG Rules — JSON to Excel

**Convert Azure Network Security Group JSON exports into clean, priority-sorted Excel reports — available as a web app, a desktop script, or a Jupyter notebook.**

[![Live App](https://img.shields.io/badge/🚀%20Live%20App-Streamlit-FF4B4B?style=for-the-badge)](https://azure-nsg-json2excel.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue?style=for-the-badge)](LICENSE)
[![Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io/)

</div>

---

## Overview

Reviewing Azure NSG security rules across environments is error-prone when everything lives in raw JSON — wildcard ports, multi-value address arrays, and default system rules all buried in nested properties. This tool parses an Azure Network Security Group (NSG) JSON export and produces a single, formatted Excel workbook with every inbound and outbound rule — custom and default — sorted by direction and priority, with wildcards replaced by the human-readable `Any`.

Available in three forms to suit any workflow: a hosted web app, a local Python script with a GUI file picker, and a Jupyter notebook for interactive exploration.

**Try it live →** [azure-nsg-json2excel.streamlit.app](https://azure-nsg-json2excel.streamlit.app/)

---

## Features

- **Three usage modes** — hosted web app, standalone Python desktop script, and Jupyter notebook
- **Inbound before Outbound** — rules are always sorted by direction, then by priority ascending
- **Custom + default rules** — merges `securityRules` and `defaultSecurityRules` into one unified table
- **Wildcard expansion** — replaces `*` with `Any` throughout ports, sources, and destinations
- **Multi-value array flattening** — expands `destinationPortRanges`, `sourceAddressPrefixes`, and `destinationAddressPrefixes` arrays into readable comma-separated strings
- **NSG metadata header** — captures NSG name, Resource Group, Location (human-readable region name), and Subscription ID
- **Formatted Excel output** — titled workbook with colour-coded headers, borders, and auto-sized columns
- **Comprehensive region mapping** — converts Azure location codes (e.g. `australiaeast`) to friendly names across 60+ regions globally
- **Live data preview** — view metadata and the full rules table in the browser before downloading (Streamlit app)

---

## Excel Output Structure

The generated workbook (`NSG_RULES` sheet) is organised into two sections:

| Section | Contents |
|---|---|
| **Metadata** | NSG name, Resource Group, Location, Subscription ID |
| **NSG RULES** | Priority · Direction · Rule Name · Port · Protocol · Source · Destination · Access · Description |

Rules are sorted **Inbound → Outbound**, then by **Priority (ascending)** within each direction.

---

## Getting Started

### Option 1 — Hosted Web App (no setup)

Visit **[azure-nsg-json2excel.streamlit.app](https://azure-nsg-json2excel.streamlit.app/)**, upload your NSG JSON file, preview the rules table, and download the Excel report.

---

### Option 2 — Run Locally (Streamlit)

**Prerequisites:** Python 3.9+

```bash
# 1. Clone the repository
git clone https://github.com/HashimsGitHub/Azure-NSG---JSON-to-Excel-converter.git
cd Azure-NSG---JSON-to-Excel-converter

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
streamlit run streamlit_app.py
```

Opens at `http://localhost:8501`.

---

### Option 3 — Desktop Script (GUI file picker)

For users who prefer working locally without a browser. Uses a Tkinter file dialog to select the input JSON and choose where to save the output.

```bash
pip install -r requirements.txt
python Azure_NSG_Rules_JSON_to_Excel_converter.py
```

A file picker will prompt for the input `.json` file, then a save dialog will prompt for the output `.xlsx` path.

---

### Option 4 — Jupyter Notebook

Open `Azure_NSG_Rules_JSON_to_Excel_coverter.ipynb` for an interactive, cell-by-cell walkthrough of the parsing and export logic — useful for customisation or learning.

```bash
pip install -r requirements.txt
jupyter notebook Azure_NSG_Rules_JSON_to_Excel_coverter.ipynb
```

---

## How to Export an Azure NSG (JSON)

Use the Azure CLI to export the JSON file this tool expects:

```bash
# Export a specific NSG by name and resource group
az network nsg show \
  --name <NSGName> \
  --resource-group <ResourceGroupName> \
  --output json > nsg.json

# Export all NSGs in a resource group
az network nsg list \
  --resource-group <ResourceGroupName> \
  --output json > all_nsgs.json

# Export all NSGs in a subscription
az network nsg list --output json > all_nsgs.json
```

Then upload or point the script at the resulting `.json` file.

> **Tip:** You can also export from the Azure Portal by navigating to your NSG → **Export template** → download the JSON, or use the **JSON view** button on the NSG overview page.

---

## Project Structure

```
Azure-NSG---JSON-to-Excel-converter/
├── streamlit_app.py                          # Streamlit web application
├── Azure_NSG_Rules_JSON_to_Excel_converter.py   # Standalone desktop script (GUI file picker)
├── Azure_NSG_Rules_JSON_to_Excel_coverter.ipynb # Jupyter notebook
├── requirements.txt                          # Python dependencies
├── .devcontainer/                            # Dev container configuration
├── .github/                                  # GitHub Actions workflows
├── .gitignore
└── LICENSE
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `streamlit` | Web application framework |
| `pandas` | DataFrame creation, sorting, and manipulation |
| `openpyxl` | Excel workbook generation and styling |

Install all with:

```bash
pip install -r requirements.txt
```

> The desktop script additionally uses `tkinter` for the GUI file picker, which is bundled with standard Python installations.

---

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -m 'Add your feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Open a Pull Request

Ideas for future improvements: multi-NSG batch processing, colour-coding Allow vs Deny rules in Excel, service tag resolution, and comparison between two NSG exports to highlight rule changes.

---

## License

Distributed under the [Apache 2.0 License](LICENSE).

---

<div align="center">

Built with ❤️ using Streamlit & Microsoft Azure

**Hashim Hilal** — Cloud Architect

</div>
