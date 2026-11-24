import openpyxl
p = r"C:\Users\adego\Downloads\WCI_Columbus Follow-up Contacts.xlsx"
wb = openpyxl.load_workbook(p, data_only=True)
print("SHEETS:", wb.sheetnames)
for ws in wb.worksheets:
    print("\n==== TAB:", repr(ws.title), "max_row:", ws.max_row, "max_col:", ws.max_column, "====")
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        print("  (empty)"); continue
    hdr_idx = 0
    for i, r in enumerate(rows):
        if any(c not in (None, "") for c in r):
            hdr_idx = i; break
    header = rows[hdr_idx]
    print("  HEADER row", hdr_idx + 1, ":")
    for j, c in enumerate(header):
        if c not in (None, ""):
            print("     col", j + 1, "=", repr(c))
    data = [r for r in rows[hdr_idx + 1:] if any(c not in (None, "") for c in r)]
    print("  DATA rows (non-empty):", len(data))
    print("  SAMPLE first 2 data rows:")
    for r in data[:2]:
        print("     ", r)
