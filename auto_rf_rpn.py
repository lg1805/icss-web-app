from flask import Flask, request, render_template, send_file
import pandas as pd
import os
import joblib
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/processed/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
MODEL_FILE_RF = 'models/random_forest.pkl'
os.makedirs('models', exist_ok=True)

# Paths for RPN file
RPN_FILE = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\RPN.xlsx"

# Load RPN Data
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()

def get_rpn_values(component):
    match = rpn_data[rpn_data["Component"] == component]
    if not match.empty:
        return match.iloc[0][["Severity (S)", "Occurrence (O)", "Detection (D)"]].tolist()
    return [2, 2, 2]  # Default values if component not found

def determine_priority(rpn):
    if rpn >= 150:
        return "High"
    elif rpn >= 80:
        return "Moderate"
    return "Low"

def train_random_forest(df):
    if not {'Severity (S)', 'Occurrence (O)', 'Detection (D)'}.issubset(df.columns):
        print("Error: Missing Severity, Occurrence, or Detection columns in training data.")
        return

    df = df.dropna(subset=['Severity (S)', 'Occurrence (O)', 'Detection (D)'])
    df['RPN'] = df['Severity (S)'] * df['Occurrence (O)'] * df['Detection (D)']
    
    X = df[['Severity (S)', 'Occurrence (O)', 'Detection (D)']]
    y = df['RPN']
    
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    model = RandomForestClassifier(n_estimators=100, random_state=42)
    model.fit(X_train, y_train)
    joblib.dump(model, MODEL_FILE_RF)
    print("Random Forest Model retrained and saved.")

def predict_rpn(severity, occurrence, detection):
    if not os.path.exists(MODEL_FILE_RF):
        return severity * occurrence * detection  # Default calculation if model doesn't exist
    
    model = joblib.load(MODEL_FILE_RF)
    return model.predict([[severity, occurrence, detection]])[0]

@app.route('/')
def index():
    return render_template('front1.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'complaint_file' not in request.files:
        return "No file uploaded", 400
    
    file = request.files['complaint_file']
    if file.filename == '':
        return "No file selected", 400
    
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    
    df = pd.read_excel(filepath)
    
    if 'Observation' not in df.columns:
        return "Error: Required column 'Observation' is missing.", 400
    
    df["Component"] = df["Observation"].apply(lambda x: "Unknown")  # Replace this with actual component extraction logic
    
    # Assign Severity, Occurrence, Detection based on RPN file
    df[["Severity (S)", "Occurrence (O)", "Detection (D)"]] = df["Component"].apply(lambda comp: pd.Series(get_rpn_values(comp)))

    # Train Random Forest on new data
    train_random_forest(df)

    # ✅ **Fixed Syntax Error Here**
    df["RPN"] = df.apply(lambda row: predict_rpn(row["Severity (S)"], row["Occurrence (O)"], row["Detection (D)"]), axis=1)
    
    df["Priority"] = df["RPN"].apply(determine_priority)
    
    processed_filepath = os.path.join(UPLOAD_FOLDER, 'processed_' + file.filename)
    df.to_excel(processed_filepath, index=False)
    
    return send_file(processed_filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
