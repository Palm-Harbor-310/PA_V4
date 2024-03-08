
import pandas as pd
import re


# Load the original Excel file from Azure
original_df = pd.read_excel("P:/Temp/2.29.2_1_table_0.xlsx")


# Identify the column with Part Numbers using regex
part_number_column = None
part_number_pattern = re.compile(r'P\d{2}-\d{3}-\d{3}')

for column in original_df.columns:
    if original_df[column].astype(str).str.match(part_number_pattern).any():
        part_number_column = column
        break

if part_number_column is None:
    raise ValueError("Part number column not found.")
else:
    original_df.rename(columns={part_number_column: 'pr_codenum'}, inplace=True)






# Load the LPC Excel file
lpc_df = pd.read_excel(
    r"C:\Users\daniel.pace\Documents\Coding\Purchasing Automation\PA_V4\Price Comparison\Last Part Cost 3.1.xlsx",
    sheet_name='All')


po_cost = lpc_df['PO Cost']







# Map PO Cost from LPC to original dataframe based on part numbers
original_df = original_df.merge(lpc_df[['pr_codenum', 'PO Cost']], on='pr_codenum', how='left')




# Save the updated DataFrame to an Excel file
original_df.to_excel('P:/Temp/2.29.2_1_table_0.xlsx', index=False)



# Assuming df is your DataFrame
# Add a new column based on the condition
original_df['Adjusted Qty'] = original_df.apply(lambda row: round(row['PO Cost'], 2) if row['pu_price'] == row['PO Cost'] else round(row['amount']/row['PO Cost'], 2), axis=1)


"P:/Temp/3.1.2_2_table_0.xlsx"
"P:/Temp/3.1_2_table_0.xlsx"
"P:/Temp/3.xlsx"