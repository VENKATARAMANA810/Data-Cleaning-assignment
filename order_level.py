import os
import pandas as pd

# === Set your folder path ===
folder_path = r"C:/Users/91798/Downloads/Cleaning of Data & Merging into single excel/Cleaning of Data & Merging into single excel/Payout Summary & Order Level Sales"

# === List to collect all order-level rows from all files ===
combined_orders = []

# === Loop through files ===
for file in os.listdir(folder_path):
    if file.endswith((".xlsx", ".xls")):
        file_path = os.path.join(folder_path, file)
        try:
            xls = pd.ExcelFile(file_path)

            # --- Parse Summary sheet ---
            summary_df = xls.parse('Summary')
            summary_df = summary_df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)

            brand_name = summary_df.loc[0, 'Unnamed: 1'] if 'Unnamed: 1' in summary_df.columns else ''
            location = summary_df.loc[1, 'Unnamed: 1'] if 'Unnamed: 1' in summary_df.columns else ''
            city = summary_df.loc[2, 'Unnamed: 1'] if 'Unnamed: 1' in summary_df.columns else ''
            res_id = summary_df.loc[3, 'Unnamed: 1'].split('-')[-1].strip() if 'Unnamed: 1' in summary_df.columns else ''
            payout_period = summary_df.loc[6, 'Unnamed: 2'] if 'Unnamed: 2' in summary_df.columns else ''
            payout_settlement_date = summary_df.loc[7, 'Unnamed: 2'] if 'Unnamed: 2' in summary_df.columns else ''

            # --- Parse Payout Breakup sheet ---
            payout_df = xls.parse('Payout Breakup').dropna(how='all').reset_index(drop=True)

            delivered_orders = cancelled_orders = total_orders = particulars = ''
            for idx, row in payout_df.iterrows():
                row_val = str(row[0])
                if 'Delivered Orders' in row_val:
                    delivered_orders = row[1]
                if 'Cancelled Orders' in row_val:
                    cancelled_orders = row[1]
                if 'Total Orders' in row_val:
                    total_orders = row[1]
                if 'Particulars' in row_val:
                    particulars = row[1]

            # --- Parse Order Level sheet ---
            order_df = xls.parse('Order Level').dropna(how='all').reset_index(drop=True)

            # Append additional info columns to each row in order_df
            order_df['Brand'] = brand_name
            order_df['Res-Id'] = res_id
            order_df['Payout Period'] = payout_period
            order_df['Location'] = location
            order_df['City'] = city
            order_df['Payout Settlement Date'] = payout_settlement_date
            order_df['Delivered Orders'] = delivered_orders
            order_df['Cancelled Orders'] = cancelled_orders
            order_df['Total Orders (Breakup)'] = total_orders
            order_df['Particulars'] = particulars
            order_df['File Name'] = file

            combined_orders.append(order_df)

        except Exception as e:
            print(f"❌ Error processing {file}: {e}")

# === Combine all order-level data ===
if combined_orders:
    final_df = pd.concat(combined_orders, ignore_index=True)

    # === Save to Excel ===
    output_file = "Combined_Order_Level_With_Summary.xlsx"
    final_df.to_excel(output_file, index=False)
    print(f"\n✅ Combined Excel saved as: {output_file}")
else:
    print("⚠️ No data extracted. Please check file structures.")
