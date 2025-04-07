import pandas as pd
import numpy as np

# Load the Excel file
df = pd.read_excel("Genset_Components_Priority_Cleaned1.xlsx")

# Define severity values based on the priority assigned earlier
severity_map = {
    "High": 9,
    "Moderate": 6,
    "Low": 3
}

# Assign severity based on the given priority
df["Severity"] = df["Priority"].map(severity_map)

# Assign occurrence and detection values randomly within a reasonable range
np.random.seed(42)  # For reproducibility
df["Occurrence"] = np.random.randint(4, 10, size=len(df))  # Range: 4 to 9
df["Detection"] = np.random.randint(2, 10, size=len(df))  # Range: 2 to 9

# Calculate RPN
df["RPN"] = df["Severity"] * df["Occurrence"] * df["Detection"]

# Categorize Risk based on RPN values
def categorize_risk(rpn):
    if rpn >= 200:
        return "High"
    elif 100 <= rpn < 200:
        return "Moderate"
    else:
        return "Low"

df["Risk Category"] = df["RPN"].apply(categorize_risk)

# Save the updated file
df.to_excel("Genset_Components_Priority_Updated.xlsx", index=False)

print("Updated file saved successfully!")
