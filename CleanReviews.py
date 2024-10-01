import pandas as pd
import re
import unicodedata
import os
from datetime import datetime

def remove_emojis(text):
    emoji_pattern = re.compile("["
                               u"\U0001F600-\U0001F64F"  # emoticons
                               u"\U0001F300-\U0001F5FF"  # symbols & pictographs
                               u"\U0001F680-\U0001F6FF"  # transport & map symbols
                               u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
                               u"\U00002702-\U000027B0"
                               u"\U000024C2-\U0001F251"
                               "]+", flags=re.UNICODE)
    return emoji_pattern.sub(r'', text)

def clean_text(text):
    # Convert to string if not already
    text = str(text)

    # Remove emojis
    text = remove_emojis(text)

    # Remove non-ASCII characters
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
    # Remove newlines and extra spaces
    text = re.sub(r'\s+', ' ', text).strip()

    # Remove URLs
    text = re.sub(r'http\S+|www.\S+', '', text)

    # Remove special characters and digits
    text = re.sub(r'[^a-zA-Z\s]', '', text)

    return text

def get_latest_excel_file(folder_path):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not excel_files:
        return None
    return max(
        [os.path.join(folder_path, f) for f in excel_files],
        key=os.path.getmtime
    )

def process_excel(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Assuming the column with reviews is named 'Review Details'
    df['Review Details'] = df['Review Details'].apply(clean_text)

    # Save the cleaned data to a new Excel file
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    input_folder = "DataInput"
    output_folder = "DataOutput"

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Get the latest Excel file from the input folder
    latest_input_file = get_latest_excel_file(input_folder)

    if latest_input_file:
        input_file_name = os.path.basename(latest_input_file)
        output_file_name = f"Cleaned_{input_file_name}"

        input_file = latest_input_file
        output_file = os.path.join(output_folder, output_file_name)

        process_excel(input_file, output_file)
        print(f"CLEANING SUCCESSFUL. Cleaned file saved as {output_file}")
    else:
        print("No Excel file found in the input folder.")