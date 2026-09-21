import pandas as pd
import pathlib
import sys
import shutil
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment

sys.stdout.reconfigure(encoding='utf-8')

WORKSPACE_ROOT = pathlib.Path('d:/behaviour analysis')
REQUIRED_COLUMNS = [
    'Symbol',
    'Strategy',
    'Entry Date',
    'Exit Date',
    'Expected Return (%)',
    'Return We Get (%)',
    'Quarter'
]
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

def safe_save(workbook: Workbook, target: pathlib.Path):
    """Save workbook, handling permission errors by writing to a versioned temp file."""
    try:
        workbook.save(target)
        return True
    except PermissionError:
        # Try versioned filenames similar to earlier approach
        for i in range(2, 6):
            versioned = target.with_name(f"{target.stem}_v{i}{target.suffix}")
            try:
                workbook.save(versioned)
                print(f'⚠️ Permission error on {target.name}; saved as {versioned.name}')
                return True
            except PermissionError:
                continue
        print(f'❌ Could not save {target.name} after multiple attempts.')
        return False

def create_template_sheet(file_path: pathlib.Path):
    # Load if exists, else start new workbook
    if file_path.exists():
        try:
            wb = load_workbook(file_path)
        except Exception:
            wb = Workbook()
    else:
        wb = Workbook()
    # Ensure Summary sheet exists (create placeholder if missing)
    if 'Summary' not in wb.sheetnames:
        ws_sum = wb.create_sheet('Summary')
        ws_sum.append(["Auto-generated template sheet"])
    # Remove old Detailed_Trades if present
    if 'Detailed_Trades' in wb.sheetnames:
        del wb['Detailed_Trades']
    ws = wb.create_sheet('Detailed_Trades')
    # Header styling
    header_font = Font(bold=True, color='FFFFFFFF')
    header_fill = PatternFill(start_color='FF333399', end_color='FF333399', fill_type='solid')
    for col_idx, col_name in enumerate(REQUIRED_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    # Example rows
    for row_idx, row_data in enumerate(example_data, start=2):
        for col_idx, col_name in enumerate(REQUIRED_COLUMNS, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row_data[col_name])
    # Auto‑size columns
    for col_idx, _ in enumerate(REQUIRED_COLUMNS, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=r, column=col_idx).value)) for r in range(1, ws.max_row+1))
        ws.column_dimensions[col_letter].width = max_len + 2
    # Save safely
    safe_save(wb, file_path)

quarter_files = sorted(WORKSPACE_ROOT.glob('*_Combined_Best_Capital_Utilisation.xlsx'))
if not quarter_files:
    print('⚠️ No quarter files found.')
    sys.exit(0)

for qf in quarter_files:
    create_template_sheet(qf)

print('✅ All quarter files now contain a "Detailed_Trades" sheet with required columns.')
