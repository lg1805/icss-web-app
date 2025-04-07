from flask import Flask, request, jsonify, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FILE = 'top_complaints.xlsx'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return jsonify({'error': 'No file part'})
    
    files = request.files.getlist('files')
    all_complaints = []
    
    for file in files:
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        df = pd.read_excel(filepath)
        if 'Observation' in df.columns:
            all_complaints.extend(df['Observation'].dropna().tolist())
    
    # Count complaint occurrences
    complaint_counts = pd.Series(all_complaints).value_counts().head(10)
    result_df = pd.DataFrame({'Complaint': complaint_counts.index, 'Count': complaint_counts.values})
    result_df.to_excel(RESULT_FILE, index=False)
    
    return jsonify(result_df.to_dict(orient='records'))

@app.route('/download', methods=['GET'])
def download_result():
    return send_file(RESULT_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
