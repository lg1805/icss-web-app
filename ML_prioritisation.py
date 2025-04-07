from flask import Flask, request, render_template, send_file
import pandas as pd
import os
import joblib

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/processed/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Paths for model and keywords file
MODEL_PATH = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\ICSS Web App\random_forest_model.pkl"
KEYWORDS_PATH = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\Genset_Components_Priority_Cleaned.xlsx"

# Load trained model
def load_model():
    try:
        model_data = joblib.load(MODEL_PATH)
        if isinstance(model_data, tuple):
            model, vectorizer = model_data
        else:
            model = model_data
            vectorizer = None
        
        if not hasattr(model, 'predict'):
            raise ValueError("Loaded model does not have a 'predict' method")
        print("Model loaded successfully")
        return model, vectorizer
    except Exception as e:
        print(f"Error loading model: {e}")
        return None, None

model, vectorizer = load_model()

# Load keywords from Excel file
def load_keywords():
    try:
        df_keywords = pd.read_excel(KEYWORDS_PATH)
        if 'Priority' not in df_keywords.columns or 'Component / System' not in df_keywords.columns:
            raise ValueError("Missing required columns in the keywords file")
        
        keyword_priority = {}
        for _, row in df_keywords.iterrows():
            for keyword in str(row['Component / System']).split(','):
                keyword_priority[keyword.strip().lower()] = row['Priority']
        
        return keyword_priority
    except Exception as e:
        print(f"Error loading keywords file: {e}")
        return {}

priority_mapping = {'High': 3, 'Moderate': 2, 'Low': 1}
priority_keywords = load_keywords()

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
        
        # Segregate SPN and Non-SPN complaints
        spn_df = df[df['Observation'].str.contains('SPN', case=False, na=False)]
        non_spn_df = df[~df['Observation'].str.contains('SPN', case=False, na=False)].copy()
        
        # Assign priority based on keyword matching in Observation column
        def assign_priority(observation):
            for keyword, priority in priority_keywords.items():
                if keyword in str(observation).lower():
                    return priority
            return 'Low'
        
        if not non_spn_df.empty:
            try:
                non_spn_df['Priority'] = non_spn_df['Observation'].apply(assign_priority)
                non_spn_df['Priority Level'] = non_spn_df['Priority'].map(priority_mapping)
                non_spn_df.sort_values(by='Priority Level', ascending=False, inplace=True)
                non_spn_df.drop(columns=['Priority Level'], inplace=True)
            except Exception as e:
                return f"Error during priority assignment: {e}", 400
        
        # Save processed file
        processed_filepath = os.path.join(UPLOAD_FOLDER, 'processed_' + file.filename)
        with pd.ExcelWriter(processed_filepath) as writer:
            spn_df.to_excel(writer, sheet_name='SPN_Complaints', index=False)
            non_spn_df.to_excel(writer, sheet_name='Non_SPN_Complaints', index=False)
        
        return send_file(processed_filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
