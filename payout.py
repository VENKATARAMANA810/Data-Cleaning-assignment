import os
import pandas as pd

# Path to your folder
folder_path = r"C:/Users/91798/Downloads/Cleaning of Data & Merging into single excel/Cleaning of Data & Merging into single excel/Payout Summary & Order Level Sales"

combined_data = []

for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        file_path = os.path.join(folder_path, file)
        try:
            xls = pd.ExcelFile(file_path)

            # Summary sheet
            summary_df = xls.parse('Summary', skiprows=3, nrows=15)
            brand_name = summary_df.iloc[0, 1]
            res_id = summary_df.iloc[3, 1].split('-')[-1].strip()
            payout_period = summary_df.iloc[7, 2]
           

            # Payout Breakup sheet
            payout_df = xls.parse('Payout Breakup', skiprows=2)
            payout_df = payout_df.rename(columns={
                payout_df.columns[1]: "Particulars",
                payout_df.columns[2]: "Delivered Orders",
                payout_df.columns[3]: "Cancelled Orders",
                payout_df.columns[4]: "Total"
            })

            # Keep only rows with valid 'Particulars'
            for idx, row in payout_df.iterrows():
                if pd.notna(row.get("Particulars")) and row["Particulars"] != "Particulars":
                    combined_data.append({
                        "Sr.No": idx + 1,
                        "Particulars": row.get("Particulars"),
                        "Delivered Orders": row.get("Delivered Orders"),
                        "Cancelled Orders": row.get("Cancelled Orders"),
                        "Total": row.get("Total"),
                        "Brand": brand_name,
                        "Res-Id": res_id,
                        "Payout Period": payout_period,
                        "File Name": file
                    })

        except Exception as e:
            print(f"❌ Error in {file}: {e}")

# Convert to DataFrame
final_df = pd.DataFrame(combined_data)

# Save to Excel
if not final_df.empty:
    output_file = "Final_Combined_PayoutBreakup.xlsx"
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Combined Payout Breakup')
    print(f"✅ Data successfully saved to: {output_file}")
else:
    print("🚫 No valid data extracted. Please check file structures.")
