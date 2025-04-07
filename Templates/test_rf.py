import os
import pandas as pd
import pickle
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import TfidfVectorizer

# Ensure NLTK resources are available
nltk.download('punkt')
nltk.download('stopwords')

# ✅ File Paths (Update if needed)
model_path = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\ICSS Web App\random_forest_model.pkl"
input_file = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\2k25 Files\Mar'25_24-03.xlsx"
output_file = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\Complaint_Segregation_Mar.xlsx"

# ✅ Load the trained Random Forest model & TfidfVectorizer together
try:
    with open(model_path, "rb") as model_file:
        rf_model, vectorizer = pickle.load(model_file)  # Unpacking both model and vectorizer
    print(f"✅ Model and TF-IDF Vectorizer Loaded Successfully! Model Type: {type(rf_model)}")
except FileNotFoundError:
    print(f"❌ Error: Model file not found at {model_path}")
    exit()

# ✅ Function to preprocess text
def preprocess_text(text):
    if pd.isna(text):
        return ""  # Handle missing values
    tokens = word_tokenize(text.lower())  # Tokenize & lowercase
    filtered_tokens = [word for word in tokens if word.isalnum()]  # Remove punctuation
    stop_words = set(stopwords.words("english"))
    processed_text = " ".join([word for word in filtered_tokens if word not in stop_words])  # Remove stopwords
    return processed_text

# ✅ Load the complaints data
try:
    df = pd.read_excel(input_file)
    print(f"✅ Data Loaded Successfully! Shape: {df.shape}")
except FileNotFoundError:
    print(f"❌ Error: Complaint data file not found at {input_file}")
    exit()

# ✅ Ensure the required column exists
if "Observation" not in df.columns:
    print("❌ Error: 'Observation' column missing! Ensure the correct file is provided.")
    exit()

# ✅ Separate SPN and Non-SPN complaints
df["Observation"] = df["Observation"].astype(str).str.strip()

spn_complaints = df[df["Observation"].str.contains(r"\bSPN\b", case=False, na=False)]
non_spn_complaints = df[~df["Observation"].str.contains(r"\bSPN\b", case=False, na=False)]

print("✅ SPN and Non-SPN complaints separated successfully!")

# ✅ Preprocess Non-SPN complaints
non_spn_complaints["Processed_Observation"] = non_spn_complaints["Observation"].apply(preprocess_text)

# ✅ Convert text into a format the model expects
non_spn_features = vectorizer.transform(non_spn_complaints["Processed_Observation"])

# ✅ Ensure feature compatibility
expected_features = rf_model.n_features_in_
actual_features = non_spn_features.shape[1]

if actual_features != expected_features:
    print(f"❌ Error: Model expects {expected_features} features, but received {actual_features}.")
    exit()

# ✅ Predict priority for Non-SPN complaints
non_spn_complaints["Predicted Priority"] = rf_model.predict(non_spn_features)

# ✅ Define priority order for sorting
priority_order = {"High": 3, "Moderate": 2, "Low": 1}
non_spn_complaints["Priority_Order"] = non_spn_complaints["Predicted Priority"].map(priority_order)

# ✅ Sort complaints based on priority: High > Moderate > Low
non_spn_complaints = non_spn_complaints.sort_values(by="Priority_Order", ascending=False).drop(columns=["Priority_Order"])

# ✅ Save output to a single Excel file with two sheets
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    spn_complaints.to_excel(writer, sheet_name="SPN_Complaints", index=False)
    non_spn_complaints.to_excel(writer, sheet_name="Non_SPN_Prioritized_Complaints", index=False)

print(f"✅ Output saved successfully in '{output_file}' with two sheets!")
