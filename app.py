import pandas as pd
from collections import Counter

# ===== USER CONFIG =====
FILE_DAILY = r"daily.xlsx"
FILE_MONTHLY = r"monthly.xlsx"

# Excel column names
COL_TICKET_NO = "Ticket No."
COL_CALLER = "Caller"
COL_PROBLEM = "Description"
COL_ENGINEER = "Assigned Engineer"
COL_STATUS = "Status"
COL_ETR = "LogTime"  # Estimated Time to Resolve
COL_DATE = "Date"

def load_excel(file):
    return pd.read_excel(file)

def analyze_missing_etr(df):
    """Return ticket numbers with missing ETR."""
    missing_etr_df = df[df[COL_ETR].isna() | (df[COL_ETR].astype(str).str.strip() == "")]
    print(f"Tickets without proper ETR: {len(missing_etr_df)}")
    if not missing_etr_df.empty:
        print("Ticket numbers with missing ETR:")
        print(missing_etr_df[COL_TICKET_NO].tolist())
    else:
        print("All tickets have ETR updated.")

def analyze_daily_pending_engineer(df):
    """For each day, show the engineer with the most pending tickets."""
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors='coerce')
    if df[COL_DATE].isna().all():
        print("No valid dates found in daily.xlsx.")
        return

    days = df[COL_DATE].dt.date.unique()
    print("\nDaily Pending Ticket Analysis (Engineer with most pending tickets):")
    for day in sorted(days):
        day_df = df[(df[COL_DATE].dt.date == day) & (df[COL_STATUS].str.lower() == "pending")]
        if not day_df.empty:
            engineer_counts = Counter(day_df[COL_ENGINEER])
            top = engineer_counts.most_common(1)
            if top:
                engineer, count = top[0]
                print(f"{day}: {engineer} ({count} pending tickets)")
            else:
                print(f"{day}: No pending tickets.")
        else:
            print(f"{day}: No pending tickets.")

def analyze_monthly(file):
    """Analyze monthly tickets and problems."""
    df = load_excel(file)
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors='coerce')
    if df.empty or df[COL_DATE].isna().all():
        print("\nNo valid data in monthly.xlsx.")
        return

    # Top 10 users who raised max tickets
    user_counter = Counter(df[COL_CALLER])
    print("\nTop 10 users who raised most tickets last month:")
    for user, count in user_counter.most_common(10):
        print(f"{user} - {count} tickets")

    # Problem which caused maximum time
    # "Maximum time" means the highest total ETR for problems
    df_valid_etr = df.copy()
    df_valid_etr[COL_ETR] = pd.to_numeric(df_valid_etr[COL_ETR], errors='coerce')
    problem_time = df_valid_etr.groupby(COL_PROBLEM)[COL_ETR].sum().sort_values(ascending=False)
    if not problem_time.empty:
        max_problem = problem_time.index[0]
        max_time = problem_time.iloc[0]
        print(f"\nProblem which caused maximum total time: '{max_problem}' ({max_time} units of time)")
    else:
        print("\nNo valid ETR data to analyze problem durations.")

def main():
    # Load daily data
    daily_df = load_excel(FILE_DAILY)

    # 1. Tickets with missing ETR
    analyze_missing_etr(daily_df)

    # 2. Daily pending engineer analysis
    analyze_daily_pending_engineer(daily_df)

    # 3. If today is the last day of the month, run monthly analysis
    import datetime
    today = datetime.date.today()
    next_day = today + datetime.timedelta(days=1)
    if next_day.month != today.month:
        print("\n=== Monthly Analysis ===")
        analyze_monthly(FILE_MONTHLY)

if __name__ == "__main__":
    main()
