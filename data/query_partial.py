import sqlite3

conn = sqlite3.connect(r'c:\Users\admin\Desktop\LAYERED2\data\account_summaries.sqlite')
c = conn.cursor()

c.execute("SELECT COUNT(*), SUM(total_credited_amount) FROM account_summaries WHERE status = ?", ('PARTIAL',))
r = c.fetchone()
print(f"Total PARTIAL records: {r[0]:,}")
print(f"Total Disputed Amount (PARTIAL): Rs. {r[1]:,.2f}")

conn.close()
