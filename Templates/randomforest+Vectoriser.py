import os
import pandas as pd
import pickle
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score

# Download necessary NLP data
nltk.download('punkt')
nltk.download('stopwords')

# Define input folder
input_folder = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\All_Files"  # Change this to your actual folder path

# Ensure input folder exists
if not os.path.exists(input_folder):
    print("❌ Input folder not found!")
    exit()

# Function to preprocess text using NLP
def preprocess_text(text):
    if pd.isna(text):
        return ""  # Handle NaN values
    tokens = word_tokenize(text.lower())  # Tokenization
    filtered_tokens = [word for word in tokens if word.isalnum()]  # Remove punctuation
    stop_words = set(stopwords.words("english"))
    processed_text = " ".join([word for word in filtered_tokens if word not in stop_words])  # Remove stopwords
    return processed_text

# Collect all complaints and labels for training
data = []

for file_name in os.listdir(input_folder):
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        file_path = os.path.join(input_folder, file_name)
        print(f"📂 Processing file: {file_name}")

        try:
            # Read the Excel file
            df_dict = pd.read_excel(file_path, sheet_name=None)

            # Process 'Non-SPN Complaints' sheet
            if "Non_SPN_Complaints" in df_dict:
                df_non_spn = df_dict["Non_SPN_Complaints"].copy()

                if "Observation" in df_non_spn and "Priority" in df_non_spn:
                    df_non_spn["Processed Observation"] = df_non_spn["Observation"].apply(preprocess_text)

                    # Store the data
                    for obs, priority in zip(df_non_spn["Processed Observation"], df_non_spn["Priority"]):
                        data.append((obs, priority))

        except Exception as e:
            print(f"❌ Error processing {file_name}: {e}\n")

# Convert collected data into DataFrame
df_training = pd.DataFrame(data, columns=["Processed Observation", "Priority"])

# Check if there's enough data
if df_training.empty:
    print("❌ No valid data found for training!")
    exit()

# Train-test split
X_train, X_test, y_train, y_test = train_test_split(
    df_training["Processed Observation"], df_training["Priority"], test_size=0.2, random_state=42
)

# Convert text into numerical features using TF-IDF
vectorizer = TfidfVectorizer()
X_train_tfidf = vectorizer.fit_transform(X_train)
X_test_tfidf = vectorizer.transform(X_test)

# Train Random Forest model
rf_model = RandomForestClassifier(n_estimators=100, random_state=42)
rf_model.fit(X_train_tfidf, y_train)

# Evaluate model
y_pred = rf_model.predict(X_test_tfidf)
accuracy = accuracy_score(y_test, y_pred)
print(f"🎯 Model Accuracy: {accuracy:.2f}")
# Save the trained model and vectorizer together
with open("rf_model_.pkl", "wb") as model_file:
    pickle.dump((rf_model, vectorizer), model_file)
print("✅ Model and vectorizer saved successfully!")
