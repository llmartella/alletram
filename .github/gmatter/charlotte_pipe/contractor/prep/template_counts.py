import os
import xlwings as xw
import pandas as pd

folder_path = r'/Users/lorimartella/Documents/gmatter/charlotte_pipe/templates'

file_data = []

# Create the Excel app once, keep it hidden and silent
app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') and not filename.startswith('~$'):
        file_path = os.path.join(folder_path, filename)

        try:
            wb = app.books.open(file_path)

            # Find the first usable sheet (not "sample" or "instructions")
            valid_sheets = [
                sheet for sheet in wb.sheets
                if not ('sample' in sheet.name.lower() or 'instructions' in sheet.name.lower())
            ]

            if not valid_sheets:
                print(f"No valid sheet found in {filename}. Skipping...")
                file_data.append({'Filename': filename, 'Complete Rows (>5 columns)': 0})
                wb.close()
                continue

            sheet = valid_sheets[0]
            data = sheet.used_range.value

            complete_rows = 0

            if data and isinstance(data, list):
                for idx, row in enumerate(data, start=1):  # start=1 for natural Excel row numbers
                    if idx == 6:
                        continue  # 🛑 Skip row 6 (header row)

                    if isinstance(row, list):
                        non_empty_cells = sum(1 for cell in row if cell not in (None, '', ' '))
                        if non_empty_cells >= 4:
                            complete_rows += 1

            file_data.append({
                'Filename': filename,
                'Complete Rows (>5 columns)': complete_rows
            })

            wb.close()

        except Exception as e:
            print(f"Error reading {filename}: {e}")
            file_data.append({
                'Filename': filename,
                'Complete Rows (>5 columns)': 'Error'})
            try:
                wb.close()
            except:
                pass  # just in case wb is not open

# Quit the Excel app after all files are processed
app.quit()

# Save the results
df = pd.DataFrame(file_data)
print(df)

output_file = r'/Users/lorimartella/Documents/gmatter/charlotte_pipe/reports/templateCounts.xlsx'
df.to_excel(output_file, index=False)

print(f"\nSaved results to {output_file}")