from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import os

app = Flask(__name__)

# Ensure required libraries are installed
try:
    import openpyxl
except ImportError:
    os.system("pip install openpyxl")

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route("/")
def home():
   return render_template('front.html')

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files["file"]
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        # Perform NLP/ML classification (Dummy logic for now)
        df["Category"] = "Sample Category"
        output_file = os.path.join(UPLOAD_FOLDER, "processed_" + file.filename)
        df.to_excel(output_file, index=False, engine="openpyxl")

        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
