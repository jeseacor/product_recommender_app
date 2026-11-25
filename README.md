# Inventory Recommendation Tool

## Overview

The Inventory Recommendation Tool helps the sales team compare customer enquiries with products in the inventory file.

A customer enquiry can contain:

- A product description
- An optional part number
- An optional customer price
- An optional GST price basis

The tool returns the closest credible inventory recommendations. It can search all suppliers or only selected preferred suppliers.

The production script is:

```text
main_runtime_parameters.py
```

---

## 1. Recommended project structure

```text
inventory_recommendation_tool/
│
├── main_runtime_parameters.py
├── requirements.txt
├── README.md
├── setup_environment.bat
├── run_recommendation.bat
├── .gitignore
│
├── data/
│   ├── customer/
│   │   └── customer_enquiry.xlsx
│   │
│   └── inventory/
│       └── inventory.xlsx
│
└── output/
    └── recommendations.xlsx
```

### Important

Do not store the virtual environment inside SharePoint as a shared dependency between computers.

Each computer should create its own `.venv` folder. The project files can be stored in SharePoint, but the virtual environment should be recreated separately on every machine.

---

## 2. Prerequisites

Install Python before setting up the project.

Recommended:

- Python 3.11 or later
- Windows 10 or Windows 11
- Visual Studio Code, although it is optional
- Excel for reviewing the input and output files

During Python installation, enable:

```text
Add Python to PATH
```

Check the installation in Command Prompt:

```bat
python --version
```

Also check pip:

```bat
python -m pip --version
```

---

## 3. First-time setup

Open Command Prompt in the project folder.

An easy method is:

1. Open the project folder in File Explorer.
2. Click the folder address bar.
3. Type `cmd`.
4. Press Enter.

### Option A — One-click setup

Double-click:

```text
setup_environment.bat
```

This script will:

1. Confirm that Python is available.
2. Create `.venv`.
3. Activate the virtual environment.
4. Upgrade pip.
5. Install the packages in `requirements.txt`.

### Option B — Manual setup

Create the virtual environment:

```bat
python -m venv .venv
```

Activate it:

```bat
.venv\Scripts\activate
```

After activation, the command prompt should show:

```text
(.venv)
```

Upgrade pip:

```bat
python -m pip install --upgrade pip
```

Install the project dependencies:

```bat
python -m pip install -r requirements.txt
```

---

## 4. Required Python packages

The `requirements.txt` file contains:

```text
pandas>=2.2,<3.0
rapidfuzz>=3.0,<4.0
openpyxl>=3.1,<4.0
```

The packages are used for:

| Package | Purpose |
|---|---|
| `pandas` | Reading, filtering, processing, and exporting Excel data |
| `rapidfuzz` | Fuzzy part-number and product-description matching |
| `openpyxl` | Excel `.xlsx` reading and writing through pandas |

To reinstall the packages:

```bat
python -m pip install -r requirements.txt
```

To confirm the installed packages:

```bat
python -m pip list
```

---

## 5. Input files

### Customer enquiry file

Store the file in:

```text
data\customer\customer_enquiry.xlsx
```

Expected columns:

| Column | Required | Description |
|---|---:|---|
| `Description` | Yes | Product wording supplied by the customer |
| `PartNumber` | No | Known part number |
| `CustomerPrice` | No | Price mentioned by the customer |
| `PriceBasis` | No | `Ex GST`, `Inc GST`, or `Unknown` |

Only `Description` is mandatory.

### Inventory file

Store the file in:

```text
data\inventory\inventory.xlsx
```

Required inventory columns:

| Column | Required |
|---|---:|
| `SupPart` | Yes |
| `Description` | Yes |

Other columns used when available include:

```text
Product
Group
Supplier
Unit
CostEx
ListPriceEx
ListPriceInc
OnHandQty
ObsoleteFlag
```

The `Supplier` column is required only when preferred-supplier filtering is used.

---

## 6. Basic run command

Activate the virtual environment:

```bat
.venv\Scripts\activate
```

Run the tool:

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --output "output\recommendations.xlsx"
```

The default number of recommendations is three per customer enquiry.

The result is saved to:

```text
output\recommendations.xlsx
```

---

## 7. One-click run

After completing the first-time setup:

1. Put the customer file in `data\customer`.
2. Put the inventory file in `data\inventory`.
3. Confirm that their names match the values in `run_recommendation.bat`.
4. Double-click:

```text
run_recommendation.bat
```

The batch file uses:

```text
data\customer\customer_enquiry.xlsx
data\inventory\inventory.xlsx
output\recommendations.xlsx
```

Edit the batch file when different file names are needed.

---

## 8. Supplier filtering

### Search all suppliers

Do not provide any `--preferred-supplier` parameter:

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --output "output\recommendations.xlsx"
```

### Search one preferred supplier

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --preferred-supplier "Supplier A" ^
  --output "output\recommendations.xlsx"
```

### Search multiple preferred suppliers

Repeat the parameter:

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --preferred-supplier "Supplier A" ^
  --preferred-supplier "Supplier B" ^
  --preferred-supplier "Supplier C" ^
  --output "output\recommendations.xlsx"
```

Supplier matching is:

- Case-insensitive
- Insensitive to leading and trailing spaces
- Otherwise based on the complete supplier name

---

## 9. List valid supplier names

Before selecting suppliers, list the supplier values available in the inventory:

```bat
python main_runtime_parameters.py ^
  --inventory "data\inventory\inventory.xlsx" ^
  --list-suppliers
```

This does not run the recommendation process. It only displays the available supplier names.

Use the displayed names in `--preferred-supplier`.

---

## 10. Change the number of recommendations

The default is:

```text
3
```

Return up to five recommendations per enquiry:

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --top-n 5 ^
  --output "output\recommendations.xlsx"
```

The tool may return fewer than the requested number when it cannot find enough credible matches.

---

## 11. Combined example

This example:

- Searches two preferred suppliers
- Returns up to five recommendations
- Writes to a named output file

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --preferred-supplier "Supplier A" ^
  --preferred-supplier "Supplier B" ^
  --top-n 5 ^
  --output "output\preferred_supplier_recommendations.xlsx"
```

---

## 12. Advanced parameters

The tool supports advanced parameters, but normal sales users should keep the default values.

Available advanced parameters include:

```text
--min-description-score
--min-part-score
--fuzzy-description-weight
--word-pair-weight
--price-tie-window
--min-single-shared-word-fuzzy-score
```

Example:

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --min-description-score 78 ^
  --min-part-score 90 ^
  --fuzzy-description-weight 0.95 ^
  --word-pair-weight 0.05 ^
  --price-tie-window 5 ^
  --min-single-shared-word-fuzzy-score 88 ^
  --output "output\recommendations.xlsx"
```

The two description weights must add up to `1.0`.

Recommended production practice:

- Sales users select suppliers and the number of results.
- An administrator controls matching thresholds.
- Change one advanced parameter at a time.
- Validate changes against known good and bad enquiries before deployment.

---

## 13. Running from Visual Studio Code

1. Open the project folder in Visual Studio Code.
2. Open **Terminal > New Terminal**.
3. Activate the virtual environment:

```bat
.venv\Scripts\activate
```

4. Select the virtual-environment interpreter if prompted:

```text
.venv\Scripts\python.exe
```

5. Run the normal command:

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --output "output\recommendations.xlsx"
```

---

## 14. Setting up on another computer

Do not copy or reuse the existing `.venv` folder.

On the new computer:

1. Copy or sync the project folder.
2. Install Python.
3. Open Command Prompt in the project folder.
4. Delete `.venv` if it came from another computer.
5. Run:

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

6. Place or confirm the input files.
7. Run the recommendation command.

The `.venv` folder can always be recreated from `requirements.txt`.

---

## 15. Updating the project

When the Python script changes:

1. Replace `main_runtime_parameters.py`.
2. Keep `requirements.txt` unless a new package is required.
3. Activate the virtual environment.
4. Reinstall requirements:

```bat
python -m pip install -r requirements.txt
```

5. Run a controlled test before replacing the deployed version.

Recommended test cases include:

- Description-only enquiry
- Exact part number
- Part number inside the description
- Customer price with Ex GST
- Customer price with Inc GST
- Blank price basis
- One preferred supplier
- Multiple preferred suppliers
- No preferred supplier
- Different inventory unit rows
- Enquiry that should return no credible recommendation

---

## 16. Troubleshooting

### `python` is not recognised

Python is either not installed or was not added to PATH.

Reinstall Python and enable:

```text
Add Python to PATH
```

### `No module named pandas` or `rapidfuzz`

Activate the virtual environment and reinstall requirements:

```bat
.venv\Scripts\activate
python -m pip install -r requirements.txt
```

### The supplier filter returns no rows

List valid supplier names:

```bat
python main_runtime_parameters.py ^
  --inventory "data\inventory\inventory.xlsx" ^
  --list-suppliers
```

Then copy the correct full supplier name.

### The output file cannot be saved

Close `recommendations.xlsx` in Excel and run the tool again. Excel can lock the file while it is open.

### Input file not found

Check:

- The file name
- The folder path
- The file extension
- Quotes around paths containing spaces

Example:

```bat
--customer "C:\My Project\data\customer\customer enquiry.xlsx"
```

### The virtual environment does not activate

Confirm that this file exists:

```text
.venv\Scripts\activate.bat
```

If not, recreate the environment:

```bat
python -m venv .venv
```

### SharePoint creates duplicate or temporary files

Avoid selecting temporary Excel files beginning with:

```text
~$
```

Confirm that the SharePoint files have completed syncing before running the tool.

---

## 17. Deactivate the environment

When finished:

```bat
deactivate
```

Closing the terminal also ends the activated environment.

---

## 18. Quick start summary

First-time setup:

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Normal run:

```bat
.venv\Scripts\activate

python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --output "output\recommendations.xlsx"
```

Preferred suppliers:

```bat
python main_runtime_parameters.py ^
  --customer "data\customer\customer_enquiry.xlsx" ^
  --inventory "data\inventory\inventory.xlsx" ^
  --preferred-supplier "Supplier A" ^
  --preferred-supplier "Supplier B" ^
  --output "output\recommendations.xlsx"
```

List suppliers:

```bat
python main_runtime_parameters.py ^
  --inventory "data\inventory\inventory.xlsx" ^
  --list-suppliers
```

Starting the application:
cd /d C:\DataProjects\OPDProductRecommendationTool
.venv\Scripts\python.exe -m streamlit run app.py

