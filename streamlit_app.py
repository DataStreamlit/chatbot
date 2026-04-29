import pandas as pd
import os
from pathlib import Path

# Main folder that contains Excel files
input_folder = r"C:\Users\Hamdan\Downloads\Adda"

# New output folder
output_folder = os.path.join(input_folder, "sample_data_1000_rows")
os.makedirs(output_folder, exist_ok=True)

# Read all Excel files
excel_files = list(Path(input_folder).glob("*.xlsx")) + list(Path(input_folder).glob("*.xls"))

for file_path in excel_files:
    try:
        print(f"Reading: {file_path.name}")

        # Read all sheets in the Excel file
        sheets = pd.read_excel(file_path, sheet_name=None)

        # Create one output Excel per source file
        output_file = os.path.join(
            output_folder,
            f"{file_path.stem}.xlsx"
        )

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            for sheet_name, df in sheets.items():
                sample_df = df.head(1000)

                # Excel sheet name max length is 31
                clean_sheet_name = str(sheet_name)[:31]

                sample_df.to_excel(
                    writer,
                    sheet_name=clean_sheet_name,
                    index=False
                )

        print(f"Saved: {output_file}")

    except Exception as e:
        print(f"Error reading {file_path.name}: {e}")

print("Done.")
