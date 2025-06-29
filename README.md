# 🧮 pyRevit Cost Estimates Extension

This extension helps automate quantity takeoffs and cost estimation directly inside Autodesk Revit using pyRevit.

---

## 📌 Features

- **Amount:** Populate cost parameter (e.g., `Test_1234`) based on category
- **Generate BOQ:** Export structured cost breakdowns to Excel
- **Grand Total:** Summarize total cost across all categories
- **Update Family Cost:** Update family cost data using a CSV-based material pricing database

---

## 🛠 Installation

Clone this repo and add it as a pyRevit extension:

```bash
pyrevit extend clone costestimates https://github.com/SwanaWJ/pyrevit-CostEstimates.git
pyrevit extensions reload

Once installed, you’ll see a "Cost Estimates" tab in your Revit ribbon.

🧩 Folder Structure
CostEstimates.extension/
├── extension.yaml
├── tab/
│   └── Cost Estimates.tab/
│       ├── Amount.pushbutton/
│       ├── Generate BOQ.pushbutton/
│       ├── Grand Total.pushbutton/
│       └── Update Family Cost.pushbutton/

🧑‍💻 Author
Wachama J. Swana
Founder of Scalefullsite – Automating engineering workflows

## 📄 License

MIT License

