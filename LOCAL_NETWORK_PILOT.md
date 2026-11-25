# Local Network Pilot — Streamlit

This pilot hosts the Inventory Recommendation Tool on one Windows PC and
lets other users on the same approved internal network open it in a browser.

## Files used

```text
inventory_recommendation_tool/
├── app.py
├── main_runtime_parameters.py
├── requirements.txt
├── .streamlit/
│   └── config.toml
├── data/
│   └── inventory/
│       └── inventory.xlsx
└── output/
```

## First-time setup

Open Command Prompt in the project folder.

```bat
python -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Activation is optional because all commands below call the virtual
environment's Python executable directly.

## Add the inventory file

Copy the current inventory workbook to:

```text
data\inventory\inventory.xlsx
```

The file must include `SupPart`, `Description`, and `Supplier` when supplier
selection is required.

## Test on the host PC

```bat
.venv\Scripts\python.exe -m streamlit run app.py
```

Open:

```text
http://localhost:8501
```

## Find the host PC's internal IPv4 address

```bat
ipconfig
```

Find the active network adapter and note its IPv4 Address, for example:

```text
192.168.1.25
```

A second PC on the same routed internal network would open:

```text
http://192.168.1.25:8501
```

## Windows Firewall

Run PowerShell as Administrator and create an inbound rule:

```powershell
New-NetFirewallRule `
  -DisplayName "Inventory Recommendation Pilot" `
  -Direction Inbound `
  -Protocol TCP `
  -LocalPort 8501 `
  -Action Allow `
  -Profile Domain,Private
```

Do not enable the rule for the Public profile.

To inspect the rule:

```powershell
Get-NetFirewallRule -DisplayName "Inventory Recommendation Pilot"
```

To remove it after the pilot:

```powershell
Remove-NetFirewallRule -DisplayName "Inventory Recommendation Pilot"
```

Company policy may prevent local firewall changes. In that case, ask IT to
approve TCP port 8501 for the host PC on the Domain or Private network.

## Start the pilot

```bat
cd "C:\Path\To\inventory_recommendation_tool"
.venv\Scripts\python.exe -m streamlit run app.py
```

Keep the Command Prompt window open. Closing it stops the application.

## Stop the pilot

In the host Command Prompt press:

```text
Ctrl+C
```

## Troubleshooting

### The app works on localhost but not another PC

Check:

1. Both PCs are on the same company network or connected through approved VPN.
2. The URL uses the host PC's IPv4 address, not `localhost`.
3. Windows Firewall allows inbound TCP 8501 on Domain or Private profile.
4. A company firewall or network policy is not blocking peer-to-peer access.
5. The host PC has not entered sleep mode.

### The browser says the site is not secure

This pilot uses HTTP and has no TLS certificate. Use it only on an approved
internal network with non-sensitive test data. HTTPS and authentication
should be added before production use.

### Inventory updates are not visible

Replace `data\inventory\inventory.xlsx` and refresh the browser. Streamlit's
cache is invalidated using the inventory file's modified timestamp.

### The output cannot be generated

Confirm that:

- The customer workbook has a `Description` column.
- The inventory workbook is not corrupt.
- Preferred supplier names exist in the inventory.
- The inventory workbook is not an Excel temporary file beginning with `~$`.
