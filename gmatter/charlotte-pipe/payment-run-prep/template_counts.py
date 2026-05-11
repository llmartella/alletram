import os
import warnings
import openpyxl
import pandas as pd
from datetime import datetime

# updated to not require opening excel and Grant Access

# --- Patch openpyxl to clamp font family values instead of raising ---
# Some xlsx files use font family=34 which exceeds openpyxl's max of 14
from openpyxl.descriptors import Max
_original_max_set = Max.__set__

def _clamped_max_set(self, instance, value):
    if value is not None:
        try:
            value = self.expected_type(value)
            if value > self.max:
                value = self.max  # clamp instead of raise
        except (ValueError, TypeError):
            pass
    _original_max_set(self, instance, value)

Max.__set__ = _clamped_max_set
# --- End patch ---


# --- Paths ---
folder_path = r'/Users/lorimartella/Documents/gmatter/charlotte_pipe/templates'
excel_output_path = r'/Users/lorimartella/Documents/gmatter/charlotte_pipe/cpf_python_scripts/outputs/template_counts.xlsx'

file_data = []

for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') and not filename.startswith('~$'):
        file_path = os.path.join(folder_path, filename)

        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                wb = openpyxl.load_workbook(file_path, data_only=True, keep_links=False)

            # Find first usable sheet
            valid_sheets = [
                sheet for sheet in wb.worksheets
                if 'sample' not in sheet.title.lower() and 'instructions' not in sheet.title.lower()
            ]

            if not valid_sheets:
                print(f"No valid sheet found in {filename}. Skipping...")
                file_data.append({'file_name': "'" + filename, 'transaction_count': 0, 'sales_total': 0})
                wb.close()
                continue

            sheet = valid_sheets[0]

            # Read all rows as a list of tuples
            data = [row for row in sheet.iter_rows(values_only=True)]
            wb.close()

            if not data:
                file_data.append({'file_name': "'" + filename, 'transaction_count': 0, 'sales_total': 0})
                continue

            # Auto-detect header row by finding the row that contains 'extended_price'
            header_idx = None
            for i, row in enumerate(data):
                if row and any(str(c).lower().strip() == "extended_price" for c in row if c):
                    header_idx = i
                    break

            if header_idx is None:
                print(f"  ⚠️  Could not find 'extended_price' header in {filename} — skipping.")
                file_data.append({'file_name': "'" + filename, 'transaction_count': 0, 'sales_total': 0})
                continue

            header_row = data[header_idx]

            df = pd.DataFrame(data[header_idx + 1:], columns=header_row)

            # Count rows with ≥2 non-empty cells
            transaction_count = df.dropna(thresh=2).shape[0]

            # Sum extended_price column
            lower_cols = [str(c).lower().strip() if c else "" for c in header_row]
            col_name = header_row[lower_cols.index("extended_price")]
            sales_total = pd.to_numeric(df[col_name], errors="coerce").sum(skipna=True)

            # Prefix filename with single quote to preserve spacing in Google Sheets
            file_data.append({
                'file_name': "'" + filename,
                'transaction_count': transaction_count,
                'sales_total': sales_total
            })

        except Exception as e:
            print(f"Error reading {filename}: {e}")
            file_data.append({
                'file_name': "'" + filename,
                'transaction_count': None,
                'sales_total': None
            })

# --- Convert results to DataFrame ---
df_out = pd.DataFrame(file_data)

print(df_out)

# --- Save to Excel ---
try:
    os.makedirs(os.path.dirname(excel_output_path), exist_ok=True)
    df_out.to_excel(excel_output_path, index=False)
    print(f"✅ Saved results to Excel: {excel_output_path}")
except Exception as e:
    print(f"⚠️ Error saving Excel: {e}")
