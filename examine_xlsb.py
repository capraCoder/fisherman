#!/usr/bin/env python3
"""
Script to examine the .xlsb file using pyxlsb
"""
from pyxlsb import open_workbook

print("Opening temiz40b.xlsb...")
with open_workbook('/home/user/fisherman/temiz40b.xlsb') as wb:
    # Get worksheet M-01
    sheet_names = wb.sheets
    print(f"Available sheets: {sheet_names}\n")

    with wb.get_sheet('M-01') as ws:
        print("--- Reading M-01 worksheet ---\n")

        # Read all rows and store data
        rows_data = []
        for row in ws.rows():
            rows_data.append([cell.v if cell else None for cell in row])

        print(f"Total rows: {len(rows_data)}\n")

        # Show headers (row 1)
        if rows_data and len(rows_data[0]) > 7:
            print("Column headers:")
            print(f"  A: {rows_data[0][0]}")
            print(f"  B: {rows_data[0][1]}")
            print(f"  C: {rows_data[0][2]}")
            print(f"  D: {rows_data[0][3]}")
            print(f"  E: {rows_data[0][4]}")
            print(f"  F: {rows_data[0][5]}")
            print(f"  G: {rows_data[0][6]}")
            print(f"  H: {rows_data[0][7]}")

        # Show sample data (rows 2-10)
        print("\n--- Sample data (rows 2-10) ---")
        for idx in range(1, min(10, len(rows_data))):
            row = rows_data[idx]
            c_val = row[2] if len(row) > 2 else None
            d_val = row[3] if len(row) > 3 else None
            print(f"Row {idx+1}: C={c_val}, D={d_val}")

        # Show row 86 data
        if len(rows_data) > 85:
            row86 = rows_data[85]
            print(f"\n--- Row 86 data ---")
            print(f"  C86: {row86[2] if len(row86) > 2 else None}")
            print(f"  D86: {row86[3] if len(row86) > 3 else None}")
            print(f"  H86: {row86[7] if len(row86) > 7 else None}")

        # Show context around row 86
        print(f"\n--- Context around row 86 (rows 80-92) ---")
        for idx in range(79, min(92, len(rows_data))):
            row = rows_data[idx]
            c_val = row[2] if len(row) > 2 else None
            d_val = row[3] if len(row) > 3 else None
            h_val = row[7] if len(row) > 7 else None
            marker = " <-- ROW 86" if idx == 85 else ""
            print(f"Row {idx+1}: C={c_val}, D={d_val}, H={h_val}{marker}")

        # Find LTQ rows
        print("\n--- Searching for 'LTQ' in column C ---")
        ltq_rows = []
        for idx, row in enumerate(rows_data):
            if len(row) > 2 and row[2]:
                c_val_str = str(row[2]).strip()
                if c_val_str == 'LTQ':
                    d_val = row[3] if len(row) > 3 else None
                    ltq_rows.append((idx+1, c_val_str, d_val))
                    if len(ltq_rows) <= 20:  # Show first 20
                        print(f"  Row {idx+1}: C={c_val_str}, D={d_val}")

        if len(ltq_rows) > 20:
            print(f"  ... and {len(ltq_rows) - 20} more rows")

        print(f"\nTotal rows with 'LTQ' in column C: {len(ltq_rows)}")

        # Calculate what the sum should be
        ltq_sum = 0
        for row_num, c_val, d_val in ltq_rows:
            if d_val is not None:
                try:
                    ltq_sum += float(d_val)
                except (ValueError, TypeError):
                    print(f"  Warning: Could not convert D{row_num}='{d_val}' to number")

        print(f"Expected sum of column D where C='LTQ': {ltq_sum}")

        print("\n" + "="*60)
        print("NOTE: pyxlsb can only read cell VALUES, not FORMULAS.")
        print("The value in H86 shown above is the CALCULATED result.")
        print("To see the actual formula, we need to check the .xlsx version.")
        print("="*60)
