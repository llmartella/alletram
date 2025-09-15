# before running this clear out /Users/lorimartella/Documents/gmatter/charlotte_pipe/templates and add new files for current pay run
# this script will create a new file containing the file name, count of records in each file and total sales amount

import os
import xlwings as xw
import pandas as pd

folder_path = r'/Users/lorimartella/Documents/gmatter/charlotte_pipe/templates'
output_file = r'/Users/lorimartella/Documents/gmatter/charlotte_pipe/reports/templateCounts.xlsx'

file_data = []

# Create one hidden Excel instance for all files
app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') and not filename.startswith('~$'):
        file_path = os.path.join(folder_path, filename)

        try:
            wb = app.books.open(file_path)

            # Find first usable sheet
            valid_sheets = [
                sheet for sheet in wb.sheets
                if 'sample' not in sheet.name.lower() and 'instructions' not in sheet.name.lower()
            ]

            if not valid_sheets:
                print(f"No valid sheet found in {filename}. Skipping...")
                file_data.append({'Filename': filename, 'Complete Rows (>5 columns)': 0, 'Sales': 0})
                wb.close()
                continue

            sheet = valid_sheets[0]
            data = sheet.used_range.value

            if not data or not isinstance(data, list):
                file_data.append({'Filename': filename, 'Complete Rows (>5 columns)': 0, 'Sales': 0})
                wb.close()
                continue

            header_row = data[5] if len(data) > 5 else None  # Excel row 6 = index 5
            df = pd.DataFrame(data[6:], columns=header_row)  # everything after header row

            # Count rows with ≥4 non-empty cells
            complete_rows = df.dropna(thresh=4).shape[0]

            # Sum Sales
            if header_row:
                lower_cols = [str(c).lower().strip() if c else "" for c in header_row]
                if "extended_price" in lower_cols:
                    col_name = header_row[lower_cols.index("extended_price")]
                    sales_sum = pd.to_numeric(df[col_name], errors="coerce").sum(skipna=True)
                else:
                    sales_sum = 0
            else:
                sales_sum = 0

            file_data.append({
                'Filename': filename,
                'Complete Rows (>5 columns)': complete_rows,
                'Sales': sales_sum
            })

            wb.close()

        except Exception as e:
            print(f"Error reading {filename}: {e}")
            file_data.append({
                'Filename': filename,
                'Complete Rows (>5 columns)': 'Error',
                'Sales': 'Error'
            })
            try:
                wb.close()
            except:
                pass

# Quit Excel after processing all files
app.quit()

# Save results
df = pd.DataFrame(file_data)
print(df)

df.to_excel(output_file, index=False)
print(f"\nSaved results to {output_file}")
