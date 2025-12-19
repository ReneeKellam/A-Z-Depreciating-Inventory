# Inventory Depreciation Script, Compares current and past inventory files to find common items
# and then marks them as inactive for depreciation purposes.
# Designed to run on a fresh instance of Python with no pre-installed packages

# Editable Variables
invcurrent_loc = r"C:\Users\azradmin\Downloads\Invcurrent.csv"
invpast_loc = r"C:\Users\azradmin\Downloads\Invpast.xlsx"
export_loc = r"C:\Users\azradmin\Downloads\Common_Items.csv"


# Python library management to ensure required packages are installed
import sys
import subprocess

def install_package(package):
    #Install a package using pip
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"Successfully installed {package}")
    except subprocess.CalledProcessError:
        print(f"Failed to install {package}")
        sys.exit(1)

# Try to import required packages, install if not available
try:
    import pandas as pd
    print("pandas imported successfully")
except ImportError:
    print("pandas not found, installing...")
    install_package("pandas")
    import pandas as pd
    print("pandas imported successfully")

# datetime is part of Python standard library, but let's handle it safely
try:
    import datetime
    print("datetime imported successfully")
except ImportError:
    # This should never happen as datetime is built-in, but just in case
    print("datetime module not found rare error.")
    sys.exit(1)

# Start of the main script
print("\nStarting inventory comparison script...")
# Try different encodings for the CSV file
try:
    invcurrent = pd.read_csv(invcurrent_loc, 
                          low_memory=False, encoding='latin-1')
except UnicodeDecodeError:
    try:
        invcurrent = pd.read_csv(invcurrent_loc, 
                              low_memory=False, encoding='cp1252')
    except UnicodeDecodeError:
        invcurrent = pd.read_csv(invcurrent_loc, 
                              low_memory=False, encoding='utf-8', errors='ignore')

invpast = pd.read_excel(invpast_loc)


# Remove Inactive items from both dataframes
print("\nRemoving inactive items...")
invcurrent = invcurrent[invcurrent['Inactive'] != True]
print("successfully removed inactive from invcurrent")
invpast = invpast[invpast['Active?'] != 'Inactive']
print("successfully removed inactive from invpast")
print()

# Debug: Check data types and samples
# print("\nDebug Info:")
# print(f"invcurrent Item ID type: {invcurrent['Item ID'].dtype}")
# print(f"invpast Item ID type: {invpast['Item ID'].dtype}")
# print(f"Sample invcurrent Item IDs: {invcurrent['Item ID'].head(5).tolist()}")
# print(f"Sample invpast Item IDs: {invpast['Item ID'].head(5).tolist()}")

# Clean the Item IDs (remove whitespace and ensure same type)
invcurrent['Item ID'] = invcurrent['Item ID'].astype(str).str.strip()
invpast['Item ID'] = invpast['Item ID'].astype(str).str.strip()

# Common Item IDs
common_items = invcurrent[invcurrent['Item ID'].isin(invpast['Item ID'])]

# Verify: Check if any Item ID from result is NOT in invpast
verification = common_items['Item ID'].isin(invpast['Item ID'])
if not verification.all():
    print("\nERROR: Found items in result that are not in invpast!")
    print(common_items[~verification]['Item ID'].tolist())
else:
    print("\nVERIFIED: All items in result exist in invpast")

# Removing new equipment items, assemblies, and services (Item Class must be 0)
common_items = common_items[common_items['Item Class'] == 0]

print("\nComparison Results:")
print(f"Total items in invcurrent: {len(invcurrent)}")
print(f"Total items in invpast: {len(invpast)}")
print(f"Common items found: {len(common_items)}")

# Preparing export dataframe
common_items_export = common_items[['Item ID', "Inactive", 'Description for Sales', "Part Number"]].reset_index(drop=True)

# Adding in depreciated inventory information to the export
for index, row in common_items_export.iterrows():
    Description = str(row['Description for Sales'])
    part_number = row['Part Number']

    mmYYYY = datetime.datetime.now().strftime("%m%Y")

    Description += " - DEP INV"
    part_number = f"DEPINV {mmYYYY}-10YRS"

    # check if description is less than or equal to 160 characters
    while len(Description) > 160:
        print(f"\n\n{index+1}/{len(common_items_export)}: Description of {row['Item ID']} exceeds 160 characters:")
        print(Description)
        print("Please shorten the description manually.")
        new_description = input("Enter new description: ")
        Description = new_description

    common_items_export.at[index, 'Part Number'] = part_number
    common_items_export.at[index, 'Description for Sales'] = Description.upper()
    common_items_export.at[index, 'Inactive'] = "TRUE"

# Exporting to CSV
common_items_export.to_csv(export_loc, index=False, encoding='utf-8-sig')
print("\nCommon items exported successfully.\n")
