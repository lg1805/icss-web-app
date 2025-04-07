import pandas as pd
import re

# Paths
COMPLAINTS_FILE = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\All-files\Apr'24.xlsx"  # Update this path
RPN_FILE = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\RPN.xlsx'  # Update this path
UPDATED_COMPLAINTS_FILE = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\RPN-assigned\rpn-assign-apr24.xlsx'

# Load complaints dataset
df = pd.read_excel(COMPLAINTS_FILE)

# Load RPN dataset
rpn_data = pd.read_excel(RPN_FILE)

# Ensure required columns exist
if 'Observation' not in df.columns:
    raise ValueError("Dataset must contain 'Observation' column!")

if 'Component' not in rpn_data.columns or 'Severity (S)' not in rpn_data.columns:
    raise ValueError("RPN file must contain 'Component', 'Severity (S)', 'Occurrence (O)', and 'Detection (D)' columns!")

# Convert to lowercase for case-insensitive matching
rpn_data['Component'] = rpn_data['Component'].str.lower()
df['Observation'] = df['Observation'].astype(str).str.lower()

# Function to extract component from Observation
def extract_component(observation):
    """Try to match a component from the RPN file. If no match, assign a low RPN value."""
    for component in rpn_data['Component']:
        if re.search(r'\b' + re.escape(component) + r'\b', observation):
            return component  # Return matched component
    return "unknown"  # Default value for unmatched components

# Apply function to create a new "Component" column
df["Component"] = df["Observation"].apply(extract_component)

# Merge RPN values based on extracted Component
df = df.merge(rpn_data[['Component', 'Severity (S)', 'Occurrence (O)', 'Detection (D)']], on='Component', how='left')

# Assign low RPN values for unmatched components
df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']] = df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']].fillna({
    'Severity (S)': 1, 'Occurrence (O)': 1, 'Detection (D)': 1
})

# Convert to integer type
df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']] = df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']].astype(int)

# Calculate RPN (Risk Priority Number)
df['RPN'] = df['Severity (S)'] * df['Occurrence (O)'] * df['Detection (D)']

# Save updated dataset
df.to_excel(UPDATED_COMPLAINTS_FILE, index=False)
print(f"Updated complaints with RPN values saved to {UPDATED_COMPLAINTS_FILE}")
