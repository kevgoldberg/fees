from flask import Flask
import pytest
import pandas as pd
from io import BytesIO
from app.workbook_processor import process_workbook

def test_process_workbook_valid_data():
    # Create a mock Excel file in memory
    excel_data = {
        'Account Number': [1, 2],
        'Household Full Name': ['John Doe', 'Jane Smith'],
        'Balance Due': [100, 200],
        'Pay Method': ['Check', 'Credit'],
        'Fund Family': ['Fidelity', 'Schwab'],
        'Annualized Fee %': [0.03, 0.02],
        'Account Type': ['Qualified', 'Non-Qualified'],
    }
    df = pd.DataFrame(excel_data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Cash Funding Week 1', index=False)
        df.to_excel(writer, sheet_name='Report Batch Week 1', index=False)
        output.seek(0)

    # Process the workbook
    result = process_workbook(output, week=1)

    # Load the result into a DataFrame for assertions
    result_wb = load_workbook(result)
    ws_check = result_wb['Check Fee Sweep']
    ws_client = result_wb['Client Fee Sweep']

    # Check the Check Fee Sweep pivot table
    assert ws_check.max_row > 1  # Ensure there is data
    assert ws_check['A2'].value == 'John Doe'  # Check first household name

    # Check the Client Fee Sweep pivot table
    assert ws_client.max_row > 1  # Ensure there is data
    assert ws_client['A2'].value == 'Jane Smith'  # Check first household name

def test_process_workbook_empty_data():
    # Create a mock empty Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame().to_excel(writer, sheet_name='Cash Funding Week 1', index=False)
        pd.DataFrame().to_excel(writer, sheet_name='Report Batch Week 1', index=False)
        output.seek(0)

    # Process the workbook
    result = process_workbook(output, week=1)

    # Load the result into a DataFrame for assertions
    result_wb = load_workbook(result)
    assert 'Check Fee Sweep' not in result_wb.sheetnames  # Ensure no sheets created for empty data
    assert 'Client Fee Sweep' not in result_wb.sheetnames  # Ensure no sheets created for empty data