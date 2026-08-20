import sqlite3

conn = sqlite3.connect(r'c:\Users\admin\Desktop\LAYERED2\data\account_summaries.sqlite')
cursor = conn.cursor()

with open(r'c:\Users\admin\Desktop\LAYERED2\data\db_info.txt', 'w', encoding='utf-8') as f:
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = cursor.fetchall()
    f.write("=== TABLES ===\n")
    for t in tables:
        f.write(t[0] + "\n")

    f.write("\n=== COLUMNS IN account_summaries ===\n")
    cursor.execute("PRAGMA table_info(account_summaries)")
    cols = cursor.fetchall()
    for c in cols:
        f.write(f"  {c[1]} ({c[2]})\n")

    cursor.execute("SELECT COUNT(*) FROM account_summaries")
    f.write(f"\nTotal rows: {cursor.fetchone()[0]}\n")

    f.write("\n=== SAMPLE DATA (first 3 rows) ===\n")
    cursor.execute("SELECT * FROM account_summaries LIMIT 3")
    col_names = [desc[0] for desc in cursor.description]
    f.write("Columns: " + str(col_names) + "\n")
    rows = cursor.fetchall()
    for r in rows:
        f.write(str(r) + "\n")

    # Check for bank column
    f.write("\n=== DISTINCT BANK NAMES ===\n")
    bank_col = None
    for col in col_names:
        if 'bank' in col.lower():
            bank_col = col
            break
    
    if bank_col:
        f.write(f"Bank column found: {bank_col}\n")
        cursor.execute(f"SELECT DISTINCT [{bank_col}] FROM account_summaries")
        banks = cursor.fetchall()
        for b in banks:
            f.write(f"  {b[0]}\n")
        f.write(f"\nTotal distinct banks: {len(banks)}\n")
    else:
        f.write("No bank column found in account_summaries\n")
        f.write("Looking in other tables...\n")
        for t in tables:
            tname = t[0]
            cursor.execute(f"PRAGMA table_info({tname})")
            tcols = cursor.fetchall()
            for c in tcols:
                if 'bank' in c[1].lower():
                    f.write(f"  Found '{c[1]}' in table '{tname}'\n")

    # Check bank_summaries table
    f.write("\n=== COLUMNS IN bank_summaries ===\n")
    cursor.execute("PRAGMA table_info(bank_summaries)")
    cols2 = cursor.fetchall()
    for c in cols2:
        f.write(f"  {c[1]} ({c[2]})\n")
    
    cursor.execute("SELECT * FROM bank_summaries LIMIT 3")
    col_names2 = [desc[0] for desc in cursor.description]
    f.write("\nSample data:\n")
    f.write("Columns: " + str(col_names2) + "\n")
    rows2 = cursor.fetchall()
    for r in rows2:
        f.write(str(r) + "\n")

conn.close()
print("Done - check db_info.txt")
