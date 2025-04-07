import pandas as pd
import torch
import os
from sentence_transformers import SentenceTransformer, util
import re
from sklearn.ensemble import RandomForestClassifier
import joblib

# Paths
COMPLAINTS_FOLDER = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\RPN-assigned"  # Folder containing complaint files
RPN_FILE = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\RPN.xlsx'  # Update this path
UPDATED_COMPLAINTS_FOLDER = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\RPN-assigned'  # Folder to save processed files
TRAINED_MODEL_PATH = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\Models\random_forest.pkl'  # Path to save trained model
VECTORIZER_PATH = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\Models\tfidf_vectorizer.pkl'  # Path to save vectorizer

# Ensure necessary directories exist
os.makedirs(os.path.dirname(TRAINED_MODEL_PATH), exist_ok=True)
os.makedirs(os.path.dirname(UPDATED_COMPLAINTS_FOLDER), exist_ok=True)

# Load RPN dataset
rpn_data = pd.read_excel(RPN_FILE)

# Ensure required columns exist
required_rpn_columns = {'Component', 'Severity (S)', 'Occurrence (O)', 'Detection (D)'}
if not required_rpn_columns.issubset(rpn_data.columns):
    raise ValueError(f"RPN file must contain {required_rpn_columns} columns!")

# Load BERT model for sentence similarity
bert_model = SentenceTransformer('all-MiniLM-L6-v2')

# Generate embeddings for RPN components
rpn_data['Component'] = rpn_data['Component'].str.lower()
component_embeddings = bert_model.encode(rpn_data['Component'].tolist(), convert_to_tensor=True)

def extract_component(observation):
    """Find the best matching component using keyword matching first, then BERT embeddings."""
    observation = observation.lower()
    for component in rpn_data['Component']:
        if re.search(rf'\b{re.escape(component)}\b', observation):
            return component  # Return matched component
    
    # If no keyword match, use BERT similarity
    obs_embedding = bert_model.encode(observation, convert_to_tensor=True)
    similarities = util.pytorch_cos_sim(obs_embedding, component_embeddings)[0]
    best_match_idx = similarities.argmax().item()
    best_match_score = similarities[best_match_idx].item()
    
    # If no strong match, assign low RPN values
    return rpn_data['Component'].iloc[best_match_idx] if best_match_score > 0.5 else "unknown"

# Process all files in the complaints folder
all_data = []
for file in os.listdir(COMPLAINTS_FOLDER):
    if file.endswith(".xlsx"):
        file_path = os.path.join(COMPLAINTS_FOLDER, file)
        df = pd.read_excel(file_path)
        
        if 'Observation' not in df.columns:
            print(f"Skipping {file} - Missing 'Observation' column!")
            continue
        
        df['Component'] = df['Observation'].astype(str).apply(extract_component)
        df = df.merge(rpn_data[['Component', 'Severity (S)', 'Occurrence (O)', 'Detection (D)']], on='Component', how='left', suffixes=(None, '_rpn'))
        
        rpn_columns = ['Severity (S)', 'Occurrence (O)', 'Detection (D)']
        df[rpn_columns] = df[rpn_columns].fillna({'Severity (S)': 2, 'Occurrence (O)': 2, 'Detection (D)': 2})
        
        df[rpn_columns] = df[rpn_columns].astype(int)
        df['RPN'] = df['Severity (S)'] * df['Occurrence (O)'] * df['Detection (D)']
        
        updated_file_path = os.path.join(UPDATED_COMPLAINTS_FOLDER, f"rpn-assigned-{file}")
        df.to_excel(updated_file_path, index=False)
        print(f"Processed and saved: {updated_file_path}")
        
        all_data.append(df)

# Combine all processed data for training
if all_data:
    combined_df = pd.concat(all_data, ignore_index=True)
    X = combined_df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']]
    y = combined_df['RPN']
    
    # Train Random Forest model
    rf_model = RandomForestClassifier(n_estimators=100, random_state=42)
    rf_model.fit(X, y)
    
    # Save trained model
    joblib.dump(rf_model, TRAINED_MODEL_PATH)
    print(f"Trained model saved at {TRAINED_MODEL_PATH}")
else:
    print("No valid complaint files found for training.")