# Excel Filter Program

This program filters an Excel file to extract email addresses and split full names into first and last names.

## Requirements

- Python 3.6 or higher
- pandas library (`pip install pandas`)
- openpyxl library (`pip install openpyxl`)

## How to Use

1. **Prepare your Excel file:**
   - Name your Excel file `MSTI.xlsx`
   - Make sure it has these exact column names:
     - `Email Address`
     - `Full Name`
   - Place the file in the same folder as the program

2. **Run the program:**
   - Open Command Prompt
   - Navigate to the folder containing the program
   - Run the program:
     ```
     python run.py
     ```

3. **What the program does:**
   - Reads your Excel file
   - Uses the last column to filter rows (keeps rows with "yes" or "y")
   - Splits full names at the first space
   - Creates a new file called `filtered_MSTI.xlsx` with:
     - Email (first column)
     - Recipient First Name (second column)
     - Recipient Last Name (third column)

## Example

If your Excel file has this data:
```
Email Address    Full Name    Status
test@email.com   John Smith   yes
```

The program will create a new file with:
```
Email           Recipient First Name    Recipient Last Name
test@email.com  John                    Smith
```

## Troubleshooting

If you get an error:
1. Make sure your Excel file is named `MSTI.xlsx`
2. Make sure the file is in the same folder as the program
3. Make sure your Excel file has columns named exactly:
   - `Email Address`
   - `Full Name`
4. Make sure the Excel file is not open in another program 