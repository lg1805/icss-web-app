from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import joblib
import os
from datetime import datetime

app = Flask(__name__)

# Load the trained ML model and vectorizer
model = joblib.load("model/complaint_classifier.pkl")
vectorizer = joblib.load("model/vectorizer.pkl")

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def classify_priority(text):
    """Predict priority using ML model."""
    return model.predict(vectorizer.transform([text]))[0]

def classify_timing(pending_hours, status):
    """Apply color coding based on pending hours."""
    if status.lower() == "resolved":
        return "Green"
    elif pending_hours < 16:
        return "Red"
    elif 16 <= pending_hours <= 20:
        return "Yellow"
    else:
        return "Blue"

@app.route("/")
def home():
    return render_template("index.html")  # Load UI

@app.route("/upload", methods=["POST"])
def upload_file():
    file = request.files["file"]

    if file.filename == "":
        return "No file selected", 400

    file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(file_path)

    df = pd.read_excel(file_path)

    if "Observation" not in df.columns or "Creation Date" not in df.columns or "Incident Status" not in df.columns:
        return "Missing required columns: 'Observation', 'Creation Date', 'Incident Status'", 400

    df["Observation"] = df["Observation"].astype(str)
    df["Predicted Priority"] = df["Observation"].apply(classify_priority)

    # Calculate pending hours
    now = datetime.now()
    df["Pending Hours"] = (now - pd.to_datetime(df["Creation Date"])).dt.total_seconds() / 3600

    # Apply color coding
    df["Time Status"] = df.apply(lambda row: classify_timing(row["Pending Hours"], row["Incident Status"]), axis=1)

    processed_file_path = os.path.join(app.config["UPLOAD_FOLDER"], "Processed_" + file.filename)
    df.to_excel(processed_file_path, index=False)

    return send_file(processed_file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
