import pandas as pd
import re

# Load the Excel file
file_path = r"C:\Users\mr\Downloads\ICORE_STMT_048205006288.xlsx"  # Update with your actual file path
df = pd.read_excel(file_path, sheet_name="ICORE_STMT_048205006288")  # Adjust sheet name if needed

# Function to extract structured data from the "Narration" column
def extract_narration_details(narration):
    if pd.isna(narration):  # Handle missing values
        return [""] * 7  # Ensure empty values instead of "N/A"

    if narration.startswith("CAM/"):  
        match = re.search(r"(\d{2}[-./]\d{2}[-./]\d{2,4})", narration)
        if match:
            txn_date = match.group(1)  # Extract the date
            narration = narration.replace(txn_date, "__DATE__")  # Temporarily replace it
        else:
            txn_date = ""
    else:
        txn_date = ""

    if narration.startswith("IMPS Chg"):  
        return [narration] + [""] * 6  # Fill remaining columns with blank values

    split_data = re.split(r"[/-]", narration)
    split_data = [item.strip() for item in split_data if item.strip()]  # Remove extra spaces

    if "__DATE__" in split_data:
        split_data[split_data.index("__DATE__")] = txn_date

    while len(split_data) < 7:
        split_data.append("")  # Fill with empty values

    return split_data[:7]  # Limit to 7 columns for consistency

# Apply extraction function to "Narration" column
df_narration_split = df[['Narration']].copy()
column_names = [f"Field {i+1}" for i in range(7)]
df_narration_split[column_names] = df['Narration'].apply(lambda x: pd.Series(extract_narration_details(x)))

# Save extracted narration details to a new Excel file
output_file = r"C:\Users\mr\Downloads\ICORE_Narration_Split.xlsx"
df_narration_split.to_excel(output_file, index=False)
print(f"? Extracted narration details saved as: {output_file}")

# Filtering based on user input
filter_field = "Field 1"  # Change this to the required field
filter_value = "John Doe"  # Change this to the required filter value

# Apply filtering if the field exists in the DataFrame
if filter_field in df_narration_split.columns:
    filtered_df = df_narration_split[df_narration_split[filter_field].str.contains(filter_value, na=False, case=False)]
    filtered_output_file = r"C:\Users\mr\Downloads\ICORE_Narration_Filtered.xlsx"
    filtered_df.to_excel(filtered_output_file, index=False)
    print(f"? Filtered data saved as: {filtered_output_file}")
else:
    print(f"?? Field '{filter_field}' not found in the extracted data.")
