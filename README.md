# Cost saving recommendations to Excel
Save cost savings recommendations to Excel file from AWS CostExplorer API. Savings Plans and Reserved Instances recommendations are parsed and added to one worksheet each. 
## Prerequisites
- Python3 and pip3
- AWS CLI setup and configured with all profiles that 
## Setup
```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
# or
pip3 install -r requirements.txt 
```
## Run -> Generate Excel file
```bash
python3 cost_savings_recommendations_to_excel.py
```
