import pandas as pd
from collections import Counter

# ===== USER CONFIG =====
# Change this to your Excel file path
file_path = r"C:\path\to\your\data.xlsx"

# Column names (update if your Excel columns are different)
col_user = "Caller"
col_problem = "Description"
col_helpdesk = "Assigned Engineer"
col_status = "Status"
col_ETR = "LogTime"  # Estimated Time to Resolve
col_date = "Date"

# ===== LOAD DATA =====
df = pd.read_excel(daily)

# Convert Date column to datetime
df[col_date] = pd.to_datetime(df[col_date], errors='coerce')

# ===== 1. Check if ETR is properly updated =====
missing_etr = df[df[col_ETR].isna() | (df[col_ETR].astype(str).str.strip() == "")]
print(f"Tickets without proper ETR: {len(missing_etr)}")
if not missing_etr.empty:
    print(missing_etr[[col_user, col_problem, col_helpdesk, col_status, col_ETR]])

# ===== 2. Most common problems =====
problem_counts = Counter(df[col_problem])
print("\nMost common problems:")
for problem, count in problem_counts.most_common(10):
    print(f"{problem} - {count} times")

# ===== 3. Helpdesk with most pending tickets =====
pending_df = df[df[col_status].str.lower() == "pending"]
helpdesk_counts = Counter(pending_df[col_helpdesk])
if helpdesk_counts:
    top_helpdesk, top_count = helpdesk_counts.most_common(1)[0]
    print(f"\nHelpdesk with most pending tickets: {top_helpdesk} ({top_count} tickets)")
else:
    print("\nNo pending tickets found.")

# ===== 4. Top 10 users who raised most tickets last month =====
if not df.empty:
    last_month = df[col_date].max().month - 1 or 12
    last_month_year = df[col_date].max().year if last_month != 12 else df[col_date].max().year - 1
    last_month_df = df[(df[col_date].dt.month == last_month) & (df[col_date].dt.year == last_month_year)]
    user_counts = Counter(last_month_df[col_user])
    print(f"\nTop 10 users who raised most tickets in {last_month}/{last_month_year}:")
    for user, count in user_counts.most_common(10):
        print(f"{user} - {count} tickets")
else:
    print("\nNo data to analyze for last month.")
