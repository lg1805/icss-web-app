import os
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords

# Download necessary NLP data
nltk.download('punkt')
nltk.download('stopwords')

# Define input and output folder
input_folder = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\All_Files"  # Change to your actual folder path
output_folder = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\cleaned_complaints"

# Ensure output folder exists
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Function to preprocess text using NLP
def preprocess_text(text):
    if pd.isna(text):
        return ""  # Handle NaN values
    tokens = word_tokenize(text.lower())  # Tokenization
    filtered_tokens = [word for word in tokens if word.isalnum()]  # Remove punctuation
    stop_words = set(stopwords.words("english"))
    processed_text = " ".join([word for word in filtered_tokens if word not in stop_words])  # Remove stopwords
    return processed_text

# Process all Excel files in the input folder
for file_name in os.listdir(input_folder):
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        file_path = os.path.join(input_folder, file_name)
        print(f"Processing file: {file_name}")

        try:
            # Read the entire Excel file
            df_dict = pd.read_excel(file_path, sheet_name=None)

            # Process the 'Non-SPN Complaints' sheet if it exists
            if "Non_SPN_Complaints" in df_dict:
                print("Processing 'Non-SPN Complaints' sheet...")
                df_non_spn = df_dict["Non_SPN_Complaints"].copy()

                if "Observation" in df_non_spn:
                    df_non_spn["Processed Observation"] = df_non_spn["Observation"].apply(preprocess_text)
                else:
                    print("⚠️ 'Observation' column missing in 'Non-SPN Complaints'!")

                # Save updated sheet
                df_dict["Non_SPN_Complaints"] = df_non_spn
            else:
                print("⚠️ 'Non-SPN Complaints' sheet not found!")

            # Save processed data to output folder
            output_file_path = os.path.join(output_folder, f"processed_{file_name}")
            with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
                for sheet_name, df in df_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"✅ Processed and saved: {output_file_path}\n")

        except Exception as e:
            print(f"❌ Error processing {file_name}: {e}\n")

print("🎯 Processing complete for all files!")
