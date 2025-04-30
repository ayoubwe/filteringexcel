import pandas as pd
import os

def check_file_exists(file_name):
    """Check if the input file exists"""
    if not os.path.exists(file_name):
        print(f"\nERROR: Could not find {file_name}")
        print("Please make sure:")
        print("1. The file is named 'MSTI.xlsx'")
        print("2. The file is in the same folder as this script")
        print("\nFiles in current folder:")
        for file in os.listdir('.'):
            print(f"- {file}")
        input("\nPress Enter to exit...")
        exit()

def read_excel_file(file_name):
    """Read the Excel file and return the DataFrame"""
    print(f"\nReading {file_name}...")
    return pd.read_excel(file_name)

def check_required_columns(df):
    """Check if required columns exist in the DataFrame"""
    required_columns = {
        "Email Address": "Email Address",
        "Full Name": "Full Name"
    }
    
    missing_columns = [col for col in required_columns.values() if col not in df.columns]
    if missing_columns:
        print("\nERROR: Could not find required columns")
        print("Please make sure your Excel file has these exact column names:")
        for col in required_columns.values():
            print(f"- {col}")
        input("\nPress Enter to exit...")
        exit()
    
    return required_columns

def filter_rows(df, filter_column):
    """Filter rows based on the filter column"""
    df[filter_column] = df[filter_column].astype(str).str.lower()
    return df[df[filter_column].isin(['yes', 'y'])]

def split_full_name(df, full_name_col):
    """Split full name into first and last name"""
    df[full_name_col] = df[full_name_col].astype(str)
    df['First Name'] = df[full_name_col].str.split(n=1).str[0]
    df['Last Name'] = df[full_name_col].str.split(n=1).str[1]
    return df

def create_output_file(df, email_col, output_file):
    """Create the output Excel file"""
    result_df = pd.DataFrame({
        'Email': df[email_col],
        'Recipient First Name': df['First Name'],
        'Recipient Last Name': df['Last Name']
    })
    
    print(f"\nSaving to {output_file}...")
    result_df.to_excel(output_file, index=False)
    return result_df

def main():
    print("Starting Excel Filter Program...")
    print("==============================")
    
    # File names
    input_file = "MSTI.xlsx"
    output_file = "filtered_MSTI.xlsx"
    
    try:
        # Check if input file exists
        check_file_exists(input_file)
        
        # Read Excel file
        df = read_excel_file(input_file)
        
        # Show all columns
        print("\nColumns in your Excel file:")
        for i, col in enumerate(df.columns, 1):
            print(f"{i}. {col}")
        
        # Check required columns
        required_columns = check_required_columns(df)
        email_col = required_columns["Email Address"]
        full_name_col = required_columns["Full Name"]
        
        print(f"\nUsing these columns:")
        print(f"1. Email: {email_col}")
        print(f"2. Full Name: {full_name_col}")
        
        # Get the last column for filtering
        last_column = df.columns[-1]
        print(f"\nUsing this column for filtering: {last_column}")
        
        # Filter rows
        filtered_df = filter_rows(df, last_column)
        print(f"\nFound {len(filtered_df)} rows with 'yes' or 'y'")
        
        # Split full names
        print("\nSplitting full names...")
        filtered_df = split_full_name(filtered_df, full_name_col)
        
        # Create output file
        result_df = create_output_file(filtered_df, email_col, output_file)
        
        print("\n✓ Successfully completed!")
        print(f"✓ Saved {len(result_df)} rows")
        print("\nThe new Excel file has these columns:")
        print("1. Email")
        print("2. Recipient First Name")
        print("3. Recipient Last Name")
        
    except Exception as e:
        print(f"\nERROR: {str(e)}")
        print("\nPlease make sure:")
        print("1. The Excel file is not open in another program")
        print("2. The file format is correct (.xlsx)")
        print("3. The Excel file has columns named 'Email Address' and 'Full Name'")
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main() 