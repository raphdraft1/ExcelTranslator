import pandas as pd
from googletrans import Translator
from tqdm import tqdm
import os
import time

translator = Translator()

# Initialize cache
cache = {}

# Safe Translate with exception handling, retries, caching, and timeout
def safe_translate(text, retries=3, timeout=5):
    if not isinstance(text, str):
        return text  # Skip non-strings

    # Check cache first
    if text in cache:
        return cache[text]

    # If not cached, attempt to translate
    for attempt in range(retries):
        try:
            # Throttle requests
            time.sleep(0.3)
            # Translate with a timeout (Googletrans API doesn't support direct timeout, so we use retry logic)
            start_time = time.time()
            translated_text = translator.translate(text, src='zh-CN', dest='en').text
            elapsed_time = time.time() - start_time

            if elapsed_time > timeout:
                raise TimeoutError(f"Translation timed out after {elapsed_time:.2f} seconds.")

            # Store result in cache
            cache[text] = translated_text
            return translated_text
        except TimeoutError as te:
            print(f"Timeout error: {te}. Retrying...")
        except Exception as e:
            print(f"Error encountered: {e}. Retrying...")
        time.sleep(1)  # Wait before retry

    # If all retries fail, return the original text
    print(f"Translation failed for: {text}")
    cache[text] = text
    return text

# Function to process translation across DataFrame
def translate_dataframe(df, selected_columns):
    translated_df = df.copy()
    for column in tqdm(selected_columns, desc="Translating Columns"):
        if df[column].dtype == 'object':
            translated_df[column] = df[column].apply(
                lambda x: safe_translate(x) if isinstance(x, str) else x
            )
    return translated_df

# Load the Excel file
file_path = r'C:\Users\yourdocumenthere'  # Adjust path to point to your file
excel_data = pd.ExcelFile(file_path)

# Display available sheets
print("Available sheets:")
for idx, sheet_name in enumerate(excel_data.sheet_names, start=1):
    print(f"{idx}. {sheet_name}")

# Ask user for the sheet number
sheet_num = int(input("\nEnter the sheet number you want to translate: ")) - 1
if sheet_num < 0 or sheet_num >= len(excel_data.sheet_names):
    print("Invalid sheet number. Exiting.")
    exit()

selected_sheet = excel_data.sheet_names[sheet_num]
print(f"Selected sheet: {selected_sheet}")

# Load the selected sheet into a DataFrame
df = excel_data.parse(selected_sheet)

# Display available columns
print("\nAvailable columns:")
for idx, column in enumerate(df.columns, start=1):
    print(f"{idx}. {column}")

# Ask user for columns to translate
column_nums = input("\nEnter the column numbers to translate (comma-separated, e.g., 1,3,5): ")
column_indices = [int(num.strip()) - 1 for num in column_nums.split(",")]
selected_columns = [df.columns[idx] for idx in column_indices if 0 <= idx < len(df.columns)]

if not selected_columns:
    print("No valid columns selected. Exiting.")
    exit()

print(f"Selected columns: {selected_columns}")

# Create output directory
output_dir = r'C:\Users\yourtranslatedsheets'  # Adjust path as needed
os.makedirs(output_dir, exist_ok=True)

start_time = time.time()

# Translate the selected columns
translated_df = translate_dataframe(df, selected_columns)

# Save the translated sheet
sheet_output_path = os.path.join(output_dir, f"{selected_sheet}_translated.xlsx")
with pd.ExcelWriter(sheet_output_path, engine='openpyxl') as writer:
    translated_df.to_excel(writer, index=False)
print(f"\nTranslated sheet saved to: {sheet_output_path}")

end_time = time.time()
print(f"Processing time: {end_time - start_time:.2f} seconds.")