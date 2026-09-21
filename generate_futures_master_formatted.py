import pandas as pd, sys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# Paths
futures_master_path = Path(r'D:/behaviour analysis/Futures_Master_Only.xlsx')
output_path = Path(r'D:/behaviour analysis/Futures_Master_Formatted.xlsx')

# Load existing futures data
try:
    df_fut = pd.read_excel(futures_master_path, engine='openpyxl')
except Exception as e:
    print(f'❌ Failed to read {futures_master_path}: {e}')
    sys.exit(1)

# Expected base columns (ensure we have at least Symbol, Date, Close)
expected_base = ['Symbol', 'Date', 'Close']
missing = [col for col in expected_base if col not in df_fut.columns]
if missing:
    print(f'Warning: Missing expected columns {missing} in the source futures master')
    # Continue with whatever is present

# Build the new DataFrame with required columns
new_columns = [
    'Stock',            # rename Symbol
    'Date',
    'Close',
    'Position',
    'Entry Date',
    'Exit Date',
    'Expected Return',
    'Return Obtained',
    'Booked P&L',
    'Quarterly Result Date'
]

# Map existing data
data = {}
# Stock name
if 'Symbol' in df_fut.columns:
    data['Stock'] = df_fut['Symbol']
elif 'Stock' in df_fut.columns:
    data['Stock'] = df_fut['Stock']
else:
    data['Stock'] = pd.Series([None]*len(df_fut))

# Date & Close
for col in ['Date', 'Close']:
    if col in df_fut.columns:
        data[col] = df_fut[col]
    else:
        data[col] = pd.Series([None]*len(df_fut))

# Add placeholder columns with NaN
for col in new_columns[3:]:
    data[col] = pd.NA

formatted_df = pd.DataFrame(data, columns=new_columns)

# Write to Excel using openpyxl writer to allow styling
try:
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        formatted_df.to_excel(writer, index=False, sheet_name='Futures_Master')
        # Apply basic styling
        workbook = writer.book
        worksheet = writer.sheets['Futures_Master']
        # Bold header row
        header_font = Font(bold=True)
        for col_idx, _ in enumerate(formatted_df.columns, start=1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
        # Freeze header row
        worksheet.freeze_panes = worksheet['A2']
        # Auto‑fit column widths (simple heuristic)
        for i, col in enumerate(formatted_df.columns, start=1):
            max_len = max(
                formatted_df[col].astype(str).map(len).max(),
                len(col)
            ) + 2
            worksheet.column_dimensions[chr(64 + i)].width = max_len
    print(f'✅ Successfully wrote formatted workbook to {output_path}')
except Exception as e:
    print(f'❌ Failed to write formatted workbook: {e}')
    sys.exit(1)
