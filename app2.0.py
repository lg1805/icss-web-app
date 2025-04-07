from flask import Flask, request, jsonify, render_template
import os
import pandas as pd

app = Flask(__name__)

# Upload folder setup
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

@app.route("/")
def home():
    return render_template("index1.0.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400
    
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    # Process the Excel file (returning first 5 rows as an example)
    df = pd.read_excel(file_path)
    processed_data = df.head().to_dict(orient="records")

    return jsonify({"message": "File uploaded successfully", "data": processed_data})

if __name__ == "__main__":
    app.run(debug=True)
