import pandas as pd
from sqlalchemy import create_engine, event
import urllib
import datetime
import time

# === CONFIGURATION ===
excel_path = r"C:\Users\ncamic\OneDrive - ServiceMaster Restoration Services\FP&A (SharePoint Online) - Documents\Ops\Headcount Work\OG Historical Headcount.xlsx"
sheet_name = "Hx V4"
server = "svrrestore.database.windows.net"
database = "Restore"
username = "sqlAdmin"   # 🔁 Replace
password = 'Et"&Zc!QSUtG'    # 🔁 Replace

start_time = time.time()
print(f"[{datetime.datetime.now()}] Script started.\n")

# === STEP 1: Load Excel ===
print("Reading Excel file...")
df = pd.read_excel(excel_path, sheet_name=sheet_name)
print(f"✔ Excel loaded: {df.shape[0]:,} rows, {df.shape[1]:,} columns")

# === STEP 2: Unpivot Date Columns ===
id_cols = [
    "Employee Number", "Job Title", "Department Code + Name", "Region",
    "Division", "Branch", "Rehire Date", "Hire Date", "Termination Date",
    "FLSA", "Status", "Client Name",
    "Pay Rate Effective Date", "Pay Rate", "Pay Frequency", "Annualized Pay"
]

print("Unpivoting data...")
df_melted = df.melt(id_vars=id_cols, var_name="Date", value_name="Worked")
print(f"✔ Unpivoted: {df_melted.shape[0]:,} rows")

# === STEP 3: Convert Date Columns ===
print("Converting date columns to datetime...")
df_melted["Date"] = pd.to_datetime(df_melted["Date"], format="%m/%d/%Y", errors="coerce")
df_melted["Rehire Date"] = pd.to_datetime(df_melted["Rehire Date"], errors="coerce")
df_melted["Hire Date"] = pd.to_datetime(df_melted["Hire Date"], errors="coerce")
df_melted["Termination Date"] = pd.to_datetime(df_melted["Termination Date"], errors="coerce")
print("✔ Date conversion complete.")

# === STEP 4: Connect to Azure SQL Server ===
print("Connecting to Azure SQL Server...")
params = urllib.parse.quote_plus(
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
    f"Encrypt=yes;"
    f"TrustServerCertificate=no;"
    f"Connection Timeout=30;"
)

engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

@event.listens_for(engine, "before_cursor_execute")
def set_fast_executemany(conn, cursor, statement, parameters, context, executemany):
    if executemany:
        cursor.fast_executemany = True

# Optional cleanup in case of prior failures
try:
    engine.dispose()
    print("✔ Engine disposed (reset).")
except Exception as e:
    print("⚠️  Engine dispose failed:", e)

# === STEP 5: Chunked Upload ===
print("Starting chunked upload...")

chunk_size = 250_000
total_rows = len(df_melted)

try:
    with engine.begin() as connection:
        for start in range(0, total_rows, chunk_size):
            end = min(start + chunk_size, total_rows)
            chunk = df_melted.iloc[start:end]
            print(f"→ Uploading rows {start:,} to {end - 1:,}...")
            chunk.to_sql("unpivoted_headcount", con=connection, if_exists="append", index=False)
        print(f"✔ Full upload complete: {total_rows:,} rows inserted.")
except Exception as e:
    print("❌ Chunked upload failed:")
    print("Error:", e)
    engine.dispose()

# === DONE ===
elapsed = round(time.time() - start_time, 2)
print(f"\n⏱ Total time: {elapsed} seconds")
print(f"[{datetime.datetime.now()}] Script finished.")
