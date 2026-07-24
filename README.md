# TATA Claims Live Dashboard

This Streamlit dashboard reads the Garantie All Claims Report and includes only
records where `TATA AIG PLAN` is `YES`.

## Dashboard KPIs

- Number of claims
- Approvals
- Rejections
- Claims with payment received
- Amount receivable
- Payment amount received
- Outstanding amount

It also includes date, state, status, and service-centre filters; status,
monthly, state, and outcome charts; claim-level details; and CSV download.

## Installation

Open PowerShell or Command Prompt in this folder:

```bash
python -m venv .venv
```

Activate the environment on Windows:

```powershell
.venv\Scripts\activate
```

Install packages:

```bash
pip install -r requirements.txt
```

## Run with manual report upload

```bash
streamlit run app.py
```

Open the URL displayed in the terminal, normally:

```text
http://localhost:8501
```

Upload the latest `.xlsx`, `.xls`, `.csv`, `.txt`, or `.tsv` All Claims Report
from the dashboard sidebar.

You can also upload the TATA Raw Dashboard Report as the second input. The
dashboard reconciles both files using both identifiers:

- TATA `Claim ID` = Garantie `TATA AIG CLAIM NO`
- TATA `Garantie Claim ID` = Garantie `Claim Number`

The reconciliation section shows matched/unmatched claims, TATA workflow and
ageing summaries, and the payment received values from the Garantie report.

## Run with an automatically refreshed report

Set the full path of the report before starting the dashboard.

PowerShell:

```powershell
$env:TATA_CLAIMS_FILE="D:\Reports\All Claims Report.xlsx"
streamlit run app.py
```

Command Prompt:

```bat
set TATA_CLAIMS_FILE=D:\Reports\All Claims Report.xlsx
streamlit run app.py
```

The default automatic refresh is 60 seconds when a configured file is used.
The interval can be changed or disabled from the dashboard sidebar.

## Expected fields

The required field is `TATA AIG PLAN`. The report should also contain:

- `Claim Number`
- `Claim Received Date`
- `My Status`
- `Substatus`
- `Approval Type`
- `Any Rejection`
- `State`
- `Service Centre Name`
- `Payment Received`
- `Amount Receivable - As per TATA Agreement`

Missing optional fields are handled without stopping the dashboard.

## TATA Raw Dashboard fields

The second report should contain:

- `Claim ID`
- `Garantie Claim ID`
- `My Status`
- `Actual Status`
- `Sub-Status`
- `State`
- `Receivable Amount`
- `Rejection Date`
- `Groups`
