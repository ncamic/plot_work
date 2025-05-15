import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_excel("Efficiency.xlsx")


# === STEP 2: Melt (Unpivot) Month Columns ===
# Assuming 'Region' is the ID column and months are the others
df_melted = df.melt(id_vars="Region", var_name="Month", value_name="Efficiency")

# === STEP 3: Plot Boxplot (Efficiency by Region) ===
plt.figure(figsize=(8, 6))  # Optional: Set figure size
sns.boxplot(x="Region", y="Efficiency", data=df_melted)
plt.title("Efficiency Distribution by Region")
plt.ylabel("Efficiency")
plt.xlabel("Region")
plt.show()