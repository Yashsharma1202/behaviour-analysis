import pandas as pd
import pathlib
import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment

sys.stdout.reconfigure(encoding='utf-8')

# Workspace directory containing the quarter Excel files
WORKSPACE_ROOT = pathlib.Path('d:/behaviour analysis')

# Columns required for the trade‑level data
REQUIRED_COLUMNS = [
    'Symbol',
    'Strategy',            # LONG or SHORT
    'Entry Date',
    'Exit Date',
    'Expected Return (%)',
    'Return We Get (%)',
    'Quarter'
]

# Helper to create a starter DataFrame with a couple of example rows
example_data = [
    {
        'Symbol': 'TCS',
        'Strategy': 'SYNTHETIC LONG',
        'Entry Date': '2026-09-01',
        'Exit Date': '2026-09-15',
        'Expected Return (%)': 2.5,
        'Return We Get (%)': 2.3,
        'Quarter': 'Q4 2025-26'
    },
    {
        'Symbol': 'HDFCBANK',
        'Strategy': 'SYNTHETIC SHORT',
        'Entry Date': '2026-09-05',
        'Exit Date': '2026-09-20',
        'Expected Return (%)': 1.8,
        'Return We Get (%)': 1.7,
        'Quarter': 'Q4 2025-26'
    }
]

def create_template_sheet(file_path: pathlib.Path):
    # Create a new workbook with a 'Detailed_Trades' sheet
    wb = load_workbook(file_path) if file_path.exists() else None
    if wb is None:
        wb = pd.ExcelWriter(file_path, engine='openpyxl')
        # Write a placeholder Summary sheet so the file looks familiar
        pd.DataFrame({"Note": ["This file was auto‑generated as a template."]}).to_excel(wb, sheet_name='Summary', index=False)
        wb.save()
        wb = load_workbook(file_path)
    # Remove any existing Detailed_Trades sheet to avoid duplication
    if 'Detailed_Trades' in wb.sheetnames:
        del wb['Detailed_Trades']
    ws = wb.create_sheet(title='Detailed_Trades')
    # Write headers with styling
    header_font = Font(bold=True, color='FFFFFFFF')
    header_fill = PatternFill(start_color='FF333399', end_color='FF333399', fill_type='solid')
    for col_idx, col_name in enumerate(REQUIRED_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    # Write example rows
    for row_idx, row_data in enumerate(example_data, start=2):
        for col_idx, col_name in enumerate(REQUIRED_COLUMNS, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row_data[col_name])
    # Auto‑size columns
    for col_idx, _ in enumerate(REQUIRED_COLUMNS, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=r, column=col_idx).value)) for r in range(1, ws.max_row+1))
        ws.column_dimensions[col_letter].width = max_len + 2
    wb.save(file_path)
    print(f'✅ Template written to {file_path.name}')

# Find all quarter Excel files (pattern *_Combined_Best_Capital_Utilisation.xlsx)
quarter_files = sorted(WORKSPACE_ROOT.glob('*_Combined_Best_Capital_Utilisation.xlsx'))
if not quarter_files:
    print('⚠️ No quarter files found in the workspace.')
    sys.exit(1)

for qf in quarter_files:
    create_template_sheet(qf)

print('✅ All quarter files now contain a "Detailed_Trades" sheet with required columns.')
