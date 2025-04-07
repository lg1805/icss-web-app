from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from openpyxl import load_workbook

app = Flask(__name__)

# Load priority mapping from the uploaded keyword file
def load_priority_keywords(priority_file):
    try:
        priority_df = pd.read_excel(priority_file, engine='openpyxl')
        priority_df.columns = priority_df.columns.str.strip()

        if 'Component / System' not in priority_df.columns or 'Priority' not in priority_df.columns:
            raise ValueError("ERROR: Missing 'Component / System' or 'Priority' column in priority file!")

        # Define priority mapping (Text -> Numeric)
        priority_mapping_text = {
            "High": 10,
            "Moderate": 5,
            "Low": 1
        }

        # Convert priority text to numbers
        priority_mapping = {
            str(k).lower(): priority_mapping_text.get(str(v).strip(), 1)  # Default to 1 (Low)
            for k, v in zip(priority_df['Component / System'], priority_df['Priority'])
        }

        return priority_mapping
    
    except Exception as e:
        print(f"Error loading priority file: {e}")
        return {}

# Assign priority based on keyword match in Observation
def assign_priority(non_spn_df, priority_mapping):
    if 'Observation' not in non_spn_df.columns:
        print("WARNING: 'Observation' column missing in complaints file!")
        non_spn_df["Priority"] = "Low"  # Default priority (Low)
        return non_spn_df

    non_spn_df["Observation"] = non_spn_df["Observation"].astype(str).str.lower()
    non_spn_df["Priority"] = "Low"  # Default priority (Low)

    # Define priority labels instead of numbers
    priority_labels = {10: "High", 5: "Moderate", 1: "Low"}

    for component, priority in priority_mapping.items():
        mask = non_spn_df["Observation"].str.contains(rf'\b{component}\b', regex=True, na=False)
        non_spn_df.loc[mask, "Priority"] = priority_labels.get(priority, "Low")  # Default to Low if no match

    # Sort by priority order: High > Moderate > Low
    priority_order = {"High": 1, "Moderate": 2, "Low": 3}
    non_spn_df["Priority_Sort"] = non_spn_df["Priority"].map(priority_order)
    non_spn_df = non_spn_df.sort_values(by="Priority_Sort", ascending=True).drop(columns=["Priority_Sort"])

    return non_spn_df

# Segregate SPN vs Non-SPN complaints
def segregate_spn_nonspn(df):
    if 'Observation' not in df.columns:
        print("ERROR: 'Observation' column missing in complaint file!")
        return None, None

    df['Observation'] = df['Observation'].astype(str).str.lower()

    spn_complaints = df[df['Observation'].str.contains(r'\bspn\b', case=False, na=False, regex=True)]
    non_spn_complaints = df.drop(spn_complaints.index)

    return spn_complaints, non_spn_complaints

@app.route('/')
def index():
    return render_template('front2.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'complaint_file' not in request.files or 'priority_file' not in request.files:
        return "Missing required files"
    
    complaint_file = request.files['complaint_file']
    priority_file = request.files['priority_file']
    
    if complaint_file.filename == '' or priority_file.filename == '':
        return "No selected file(s)"
    
    try:
        df = pd.read_excel(complaint_file, engine='openpyxl')
        df.columns = df.columns.str.strip()
        
        priority_mapping = load_priority_keywords(priority_file)
        
        spn_df, non_spn_df = segregate_spn_nonspn(df)
        
        if spn_df is None and non_spn_df is None:
            return "Invalid file: Missing 'Observation' column"
        
        if not non_spn_df.empty:
            non_spn_df = assign_priority(non_spn_df, priority_mapping)
        
        output_file = "Segregated_Complaints.xlsx"
        
        if os.path.exists(output_file):
            os.remove(output_file)
        
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
            spn_df.to_excel(writer, sheet_name='SPN_Complaints', index=False)
            non_spn_df.to_excel(writer, sheet_name='Non_SPN_Complaints', index=False)
        
        return send_file(output_file, as_attachment=True)
    
    except Exception as e:
        return f"Error processing file: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True, port=5001)

