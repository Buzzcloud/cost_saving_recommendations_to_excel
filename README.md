# Cost saving recommendations to Excel/Xlsx

Managers that take decisions around Reserved Instances and Savings Plans purchase typically prefer to get a brief with a tool that they understand - Excel. This is why we need this script.

Store your Cost Savings Recommendations from AWS CostExplorer API into an Excel file using [XlsxWriter](https://xlsxwriter.readthedocs.io/index.html).
[Savings Plans](https://boto3.amazonaws.com/v1/documentation/api/latest/reference/services/ce.html#CostExplorer.Client.get_savings_plans_purchase_recommendation) and [Reserved Instances](https://boto3.amazonaws.com/v1/documentation/api/latest/reference/services/ce.html#CostExplorer.Client.get_reservation_purchase_recommendation) recommendations are parsed and added to worksheets.


## Prerequisites

- Python3 and pip3.
- AWS CLI setup and configured with all profiles that.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
# or
pip3 install -r requirements.txt 
```

## Generate Excel file

Run the script to generate Excel file

```bash
python3 cost_savings_recommendations_to_excel.py
```

### Arguments

To see what arguments can be feed into the script.

```bash
python3 cost_savings_recommendations_to_excel.py --help
```
