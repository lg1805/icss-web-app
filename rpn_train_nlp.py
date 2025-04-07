import pandas as pd
import torch
from sentence_transformers import SentenceTransformer, util
import re

# Paths
COMPLAINTS_FILE = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\All-files\Sep'24.xlsx"  # Update this path
RPN_FILE = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\RPN.xlsx'  # Update this path
UPDATED_COMPLAINTS_FILE = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\RPN-assigned\rpn-BERT-sep24.xlsx'

# Load complaints dataset
df = pd.read_excel(COMPLAINTS_FILE)

# Load RPN dataset
rpn_data = pd.read_excel(RPN_FILE)

# Ensure required columns exist
if 'Observation' not in df.columns:
    raise ValueError("Dataset must contain 'Observation' column!")

if 'Component' not in rpn_data.columns or 'Severity (S)' not in rpn_data.columns:
    raise ValueError("RPN file must contain 'Component', 'Severity (S)', 'Occurrence (O)', and 'Detection (D)' columns!")

# Load BERT model for sentence similarity
model = SentenceTransformer('all-MiniLM-L6-v2')

# Generate embeddings for RPN components
rpn_data['Component'] = rpn_data['Component'].str.lower()
component_embeddings = model.encode(rpn_data['Component'].tolist(), convert_to_tensor=True)

def extract_component(observation):
    """Find the best matching component using keyword matching first, then BERT embeddings."""
    observation = observation.lower()
    for component in rpn_data['Component']:
        if re.search(rf'\b{re.escape(component)}\b', observation):
            return component  # Return matched component
    
    # If no keyword match, use BERT similarity
    obs_embedding = model.encode(observation, convert_to_tensor=True)
    similarities = util.pytorch_cos_sim(obs_embedding, component_embeddings)[0]
    best_match_idx = similarities.argmax().item()
    best_match_score = similarities[best_match_idx].item()
    
    # If no strong match, assign low RPN values
    return rpn_data['Component'].iloc[best_match_idx] if best_match_score > 0.5 else "unknown"

# Apply function to create a new "Component" column
df['Component'] = df['Observation'].astype(str).apply(extract_component)

# Merge RPN values based on extracted Component
df = df.merge(rpn_data[['Component', 'Severity (S)', 'Occurrence (O)', 'Detection (D)']], on='Component', how='left')

# Assign low RPN values for unmatched components
df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']] = df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']].fillna({
    'Severity (S)': 2, 'Occurrence (O)': 2, 'Detection (D)': 2
})

# Convert to integer type
df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']] = df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']].astype(int)

# Calculate RPN (Risk Priority Number)
df['RPN'] = df['Severity (S)'] * df['Occurrence (O)'] * df['Detection (D)']

# Save updated dataset
df.to_excel(UPDATED_COMPLAINTS_FILE, index=False)
print(f"Updated complaints with RPN values saved to {UPDATED_COMPLAINTS_FILE}")