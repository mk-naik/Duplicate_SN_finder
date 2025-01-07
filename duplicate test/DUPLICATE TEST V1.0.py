import pandas as pd

# Load the two Excel files
file1 = 'File1.xlsx'  # Replace with your first file name
file2 = 'File2.xlsx'  # Replace with your second file name

sheet_name1 = 'Sheet1'  # Adjust for the sheet name in the first file
sheet_name2 = 'Sheet1'  # Adjust for the sheet name in the second file

# Load the entire sheets and extract only the "BARCODE" columns
df1 = pd.read_excel(file1, sheet_name=sheet_name1)
df2 = pd.read_excel(file2, sheet_name=sheet_name2)

# Identify barcode columns by checking for the header "BARCODE"
barcode_columns1 = [col for col in df1.columns if "BARCODE" in str(col).upper()]
barcode_columns2 = [col for col in df2.columns if "BARCODE" in str(col).upper()]

# Combine all barcode columns from both files
all_barcodes1 = pd.concat([df1[col].dropna() for col in barcode_columns1], ignore_index=True)
all_barcodes2 = pd.concat([df2[col].dropna() for col in barcode_columns2], ignore_index=True)

# Find duplicates between the two files
duplicates = all_barcodes1[all_barcodes1.isin(all_barcodes2)].reset_index(drop=True)

# Check if duplicates exist
if not duplicates.empty:
    # Save duplicates to a new Excel file
    output_file = f"{file1.split('.')[0]}_Duplicates_Between_Files.xlsx"
    duplicates.to_frame(name="Duplicate Barcodes").to_excel(output_file, index=False)
    print(f"Duplicates saved to '{output_file}'")
else:
    print("No Duplicates Found")
