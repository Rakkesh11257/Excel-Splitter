**Excel Splitter**

**Description:**
This Python script splits a single Excel file into multiple sheets based on a specified number of rows per sheet. It utilizes the `openpyxl` library to work with Excel files and the `os` module for file operations.

**Usage:**
1. Ensure Python 3.x is installed on your system.
2. Install the `openpyxl` library if not already installed: `pip install openpyxl`.
3. Place the Excel file to be split on your desktop.
4. Update the `filename` variable in the script to specify the name of the Excel file.
5. Run the script. The Excel file will be split into multiple sheets, each containing a specified number of rows.

**Example:**
```python
import openpyxl
import os

def split_excel():
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    filename = os.path.join(desktop_path, 'YourExcelFile.xlsx')
    
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    asset_ids = [cell.value for cell in ws['A'] if cell.value is not None]

    num_sheets = len(asset_ids) // 400 + (len(asset_ids) % 400 > 0)

    for i in range(num_sheets):
        new_ws = wb.create_sheet(title=f'Sheet-{i+1}')
        start_row = i * 400
        end_row = min((i + 1) * 400, len(asset_ids))
        for j in range(start_row, end_row):
            new_ws.cell(row=j - start_row + 1, column=1, value=asset_ids[j])

    wb.remove(ws)
    wb.save(filename)

# Example usage
split_excel()
```

**Notes:**
- Ensure the Excel file is correctly formatted to avoid any issues during splitting.
- Review the generated sheets in the Excel file for accuracy.


**License:**
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

**Contributing:**
Contributions are welcome. Please fork the repository, make your changes, and submit a pull request.

**Acknowledgments:**
- This script utilizes the `openpyxl` library for Excel manipulation.


**Author:**
Rakkesh R

**Contact:**
rakkesh30.mbm@gmail.com

