from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Function to assign Risk Level
def assign_risk_level(rpn):
    if rpn >= 150:
        return 'High Risk'
    elif 100 <= rpn < 150:
        return 'Moderate Risk'
    else:
        return 'Low Risk'

@app.route('/')
def index():
    return render_template('front3.html')  # Integrating front3.html

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'complaint_file' not in request.files or 'rpn_file' not in request.files:
        return jsonify({'error': 'Both complaint_file and rpn_file are required'}), 400
    
    complaint_file = request.files['complaint_file']
    rpn_file = request.files['rpn_file']
    
    complaint_path = os.path.join(UPLOAD_FOLDER, complaint_file.filename)
    rpn_path = os.path.join(UPLOAD_FOLDER, rpn_file.filename)
    
    complaint_file.save(complaint_path)
    rpn_file.save(rpn_path)
    
    # Load complaint and RPN files
    df_complaints = pd.read_excel(complaint_path)
    df_rpn = pd.read_excel(rpn_path)
    
    # Ensure column names are correct
    if 'SPN' not in df_complaints.columns or 'FMI' not in df_complaints.columns or 'Observation' not in df_complaints.columns:
        return jsonify({'error': 'Complaint file must contain SPN, FMI, and Observation columns'}), 400
    if 'SPN' not in df_rpn.columns or 'FMI' not in df_rpn.columns or 'RPN' not in df_rpn.columns:
        return jsonify({'error': 'RPN file must contain SPN, FMI, and RPN columns'}), 400
    
    # Merge complaints with RPN data
    df_merged = df_complaints.merge(df_rpn, on=['SPN', 'FMI'], how='left')
    
    # Assign Risk Level
    df_merged['Risk Level'] = df_merged['RPN'].apply(lambda x: assign_risk_level(x) if pd.notna(x) else 'Unknown')
    
    # Separate into SPN and Non-SPN complaints based on 'Observation' column
    df_spn = df_merged[df_merged['Observation'].str.contains('spn', case=False, na=False)]
    df_non_spn = df_merged[~df_merged['Observation'].str.contains('spn', case=False, na=False)]
    
    # Save to a single Excel file with two sheets
    output_path = os.path.join(UPLOAD_FOLDER, 'processed_complaints.xlsx')
    with pd.ExcelWriter(output_path) as writer:
        df_spn.to_excel(writer, sheet_name='SPN Complaints', index=False)
        df_non_spn.to_excel(writer, sheet_name='Non-SPN Complaints', index=False)
    
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
