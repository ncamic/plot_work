import pandas as pd

df = pd.read_excel(r"C:\Users\ncamic\OneDrive - ServiceMaster Restoration Services\FP&A (SharePoint Online) - Documents\Ops\Headcount Work\OG Historical Headcount.xlsx", sheet_name="Hx V3")

id_cols = [
    "Employee Number", "Job Title", "Department Code + Name", "Region",
    "Division", "Branch", "Rehire Date", "Hire Date", "Termination Date",
    "FLSA", "Status", "Client Name"
]

df_melted = df.melt(
    id_vars=id_cols,
    var_name="Date",
    value_name="Worked"
)

# Convert the 'Date' column (formerly column headers) to datetime
df_melted["Date"] = pd.to_datetime(df_melted["Date"], format="%m/%d/%Y", errors="coerce")

# Optional: also convert date columns in id_vars
df_melted["Rehire Date"] = pd.to_datetime(df_melted["Rehire Date"], errors="coerce")
df_melted["Hire Date"] = pd.to_datetime(df_melted["Hire Date"], errors="coerce")
df_melted["Termination Date"] = pd.to_datetime(df_melted["Termination Date"], errors="coerce")

# Confirm result
#print(df_melted.dtypes)
#print(df_melted.head())
df_melted.to_csv(r"C:\Users\ncamic\OneDrive - ServiceMaster Restoration Services\Documents\Python\unpivoted_headcount.csv", index=False)
