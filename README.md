# Excel Data Entry Application

This Python application facilitates data entry into an Excel spreadsheet. It includes functionality to create new sheets for each month and add data entries to the respective sheets.

## Features

### 1. Create New Month Sheet
- Allows the user to create a new sheet in an existing Excel file for a specific month.
- Prompts the user to enter the month name and automatically generates a new sheet with the provided name.

### 2. Data Entry
- Enables the user to enter data for a specified number of entries into the Excel sheet corresponding to the selected month.
- For each data entry, the user can input various details such as date, bill number, place, party name, GST number, and taxable value.
- Calculates CGST, SGST, and total amount based on the taxable value.
- Data entry is validated, and errors are handled gracefully.

### 3. Exit
- Allows the user to exit the application.

## Usage

1. Run the Python script.
2. Choose from the available options:
   - Create new month: Create a new sheet in the Excel file for a specific month. Enter the month name when prompted.
   - Enter Data: Enter data for a specified number of entries into the Excel sheet corresponding to the selected month. Provide the required details for each entry.
   - Exit: Close the application.

## Dependencies

- `numpy`
- `pandas`
- `openpyxl`
- `tkinter`
- `datetime`

Ensure that these dependencies are installed in your Python environment before running the script.

## Example

Suppose you want to track sales data for different months. You can use this application to create separate sheets for each month and enter sales transactions accordingly. This helps in organizing data and analyzing sales performance over time.

Feel free to modify and integrate this application according to your requirements.
