import os
import pandas as pd
import nltk
import re
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords

# Download necessary NLP data
nltk.download('punkt')
nltk.download('stopwords')

# Define input and output folder
input_folder = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\All_Files"  # Change to your actual folder path
output_folder = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\cleaned_complaints1"

# Ensure output folder exists
os.makedirs(output_folder, exist_ok=True)

# Function to preprocess text using NLP
def preprocess_text(text):
    if pd.isna(text):
        return ""  # Handle NaN values
    tokens = word_tokenize(text.lower())  # Tokenization
    filtered_tokens = [word for word in tokens if word.isalnum()]  # Remove punctuation
    stop_words = set(stopwords.words("english"))
    processed_text = " ".join([word for word in filtered_tokens if word not in stop_words])  # Remove stopwords
    return processed_text

# Function to extract SPN and FMI from text
def extract_spn_fmi(text):
    spn_match = re.search(r'SPN\s*(\d+)', str(text))  # Find SPN number
    fmi_match = re.search(r'FMI\s*(\d+)', str(text))  # Find FMI number
    spn = spn_match.group(1) if spn_match else None
    fmi = fmi_match.group(1) if fmi_match else None
    return spn, fmi

# Process all Excel files in the input folder
for file_name in os.listdir(input_folder):
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        file_path = os.path.join(input_folder, file_name)
        print(f"Processing file: {file_name}")

        try:
            # Read the entire Excel file
            df_dict = pd.read_excel(file_path, sheet_name=None)

            # Process 'Non-SPN Complaints' sheet if it exists
            if "Non_SPN_Complaints" in df_dict:
                print("Processing 'Non_SPN_Complaints' sheet...")
                df_non_spn = df_dict["Non_SPN_Complaints"].copy()

                if "Observation" in df_non_spn:
                    df_non_spn["Processed Observation"] = df_non_spn["Observation"].apply(preprocess_text)
                else:
                    print("⚠️ 'Observation' column missing in 'Non-SPN Complaints'!")

                # Save updated sheet
                df_dict["Non_SPN_Complaints"] = df_non_spn
            else:
                print("⚠️ 'Non-SPN Complaints' sheet not found!")

            # Process 'SPN_Complaints' sheet if it exists
            if "SPN_Complaints" in df_dict:
                print("Processing 'SPN_Complaints' sheet...")
                df_spn = df_dict["SPN_Complaints"].copy()

                if "Observation" in df_spn:
                    df_spn[['Extracted SPN', 'Extracted FMI']] = df_spn["Observation"].apply(lambda x: pd.Series(extract_spn_fmi(x)))
                else:
                    print("⚠️ 'Observation' column missing in 'SPN_Complaints'!")

                # Save updated sheet
                df_dict["SPN_Complaints"] = df_spn
            else:
                print("⚠️ 'SPN_Complaints' sheet not found!")

            # Save processed data to output folder
            output_file_path = os.path.join(output_folder, f"processed_{file_name}")
            with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
                for sheet_name, df in df_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"✅ Processed and saved: {output_file_path}\n")

        except Exception as e:
            print(f"❌ Error processing {file_name}: {e}\n")

print("🎯 Processing complete for all files!")
