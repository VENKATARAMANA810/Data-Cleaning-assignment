import os
import pandas as pd

# === Set your folder path containing all Excel annexure files ===
folder_path = "C:/Users/91798/Downloads/Cleaning of Data & Merging into single excel/Cleaning of Data & Merging into single excel/Payout Summary & Order Level Sales"  # Replace with your folder path

# === Lists to hold combined results ===
all_brand_data = []
# === Loop through each Excel file in the folder ===
for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        file_path = os.path.join(folder_path, file)
        try:
            xls = pd.ExcelFile(file_path)
            summary_df = xls.parse('Summary')
            summary_df = summary_df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)

            # Extract Brand Details
            brand_details = {
                "Brand Name": summary_df.loc[0, 'Unnamed: 1'],
                "Location": summary_df.loc[1, 'Unnamed: 1'],
                "City": summary_df.loc[2, 'Unnamed: 1'],
                "Restaurant ID": summary_df.loc[3, 'Unnamed: 1'].split('-')[-1].strip(),
                "Payout Period": summary_df.loc[6, 'Unnamed: 2'],
                "Payout Settlement Date": summary_df.loc[7, 'Unnamed: 2'],
                "Total Payout": summary_df.loc[8, 'Unnamed: 2'],
                "Total Orders": summary_df.loc[9, 'Unnamed: 2'],
                "Bank UTR": summary_df.loc[10, 'Unnamed: 2'],
                "File Name": file,
            }

            all_brand_data.append(brand_details)

        except Exception as e:
            print(f"Error processing {file}: {e}")

# === Convert to DataFrames ===
brand_df = pd.DataFrame(all_brand_data)

# === Save to a combined Excel file ===
output_file = "Combined_Brand.xlsx"
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    brand_df.to_excel(writer, index=False, sheet_name='All Brand Details')

print(f"\n✅ Combined Excel saved as: {output_file}")
