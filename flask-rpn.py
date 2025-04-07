from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/processed/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Paths for RPN file
RPN_FILE = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\RPN.xlsx"

# Load RPN Data
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()

def extract_component(observation):
    """Matches observation text to a component from the RPN file."""
    if pd.notna(observation):
        for component in known_components:
            if str(component).lower() in observation.lower():
                return component
    return "Unknown"

def get_rpn_values(component):
    """Fetch Severity, Occurrence, and Detection values from the RPN file."""
    row = rpn_data[rpn_data["Component"] == component]
    if not row.empty:
        severity = int(row["Severity (S)"].values[0])
        occurrence = int(row["Occurrence (O)"].values[0])
        detection = int(row["Detection (D)"].values[0])
        return severity, occurrence, detection
    return 1, 1, 10  # Default values if component not found

def determine_priority(rpn):
    """Assign priority based on RPN value."""
    if rpn >= 200:
        return "High"
    elif rpn >= 100:
        return "Moderate"
    else:
        return "Low"

@app.route('/')
def index():
    return render_template('front1.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'complaint_file' not in request.files:
        return "No complaint_file part", 400
    
    file = request.files['complaint_file']
    if file.filename == '':
        return "No selected file", 400
    
    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        
        try:
            df = pd.read_excel(filepath)
        except Exception as e:
            return f"Error reading file: {e}", 400
        
        if 'Observation' not in df.columns:
            return "Error: Required columns are missing in the uploaded file.", 400
        
        # Extract component dynamically
        df["Component"] = df["Observation"].apply(extract_component)
        
        # Fetch Severity, Occurrence, and Detection values
        df[["Severity (S)", "Occurrence (O)", "Detection (D)"]] = df["Component"].apply(lambda comp: pd.Series(get_rpn_values(comp)))

        # Calculate RPN and assign Priority
        df["RPN"] = df["Severity (S)"] * df["Occurrence (O)"] * df["Detection (D)"]
        df["Priority"] = df["RPN"].apply(determine_priority)
        
        # Save processed file
        processed_filepath = os.path.join(UPLOAD_FOLDER, 'processed_' + file.filename)
        df.to_excel(processed_filepath, index=False)
        
        return send_file(processed_filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
