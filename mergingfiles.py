import os
import pandas as pd

# Folder containing complaint files
COMPLAINTS_FOLDER = r'D:\Lakshya\Project\ICSS_VSCode\PROJECT\All-files'  # Update this with your actual folder path

# List all Excel files in the folder
all_files = [os.path.join(COMPLAINTS_FOLDER, f) for f in os.listdir(COMPLAINTS_FOLDER) if f.endswith('.xlsx')]

if not all_files:
    print("No Excel files found in the folder!")

# Read and merge all complaint files
df_list = [pd.read_excel(file) for file in all_files]
merged_df = pd.concat(df_list, ignore_index=True)

# Save merged dataset to a new file
merged_file_path = os.path.join(COMPLAINTS_FOLDER, 'merged_all_complaints.xlsx')
merged_df.to_excel(merged_file_path, index=False)

print(f"Merged dataset saved successfully: {merged_file_path}")
