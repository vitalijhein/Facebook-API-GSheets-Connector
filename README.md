# Google Apps Script Project - Spreadsheet Data Transformation

This is a Google Apps Script (GAS) Project which aims to manipulate and transform data from a Google Spreadsheet. The script includes various functions for data manipulation, including:

- Formatting data
- Deleting rows and columns
- Adding a new column
- Concatenating date and time
- Changing string values to standardized codes (e.g., changing country names to their respective ISO 3166-1 alpha-2 codes)
- Creating a gender column from given data
- And many more...

## Usage

To use the code in this repository, you need to:

1. Open the Google Spreadsheet where you want to run these functions.
2. Click on `Extensions > Apps Script`.
3. Paste the code into the Apps Script Editor and Save it.
4. Run the `main()` function.

Please note that the script is configured to run on spreadsheets with specific columns. The functionality may break if the spreadsheet format does not match the expected format. Make sure your data includes the following columns:

- "Bestellnummer"
- "Bestelldatum/Datum"
- "Bestelldatum/Zeit"
- "Kunde/Anrede"
- "Kunde/Vorname"
- "Kunde/Name"
- "Kunde/PLZ"
- "Kunde/Ort"
- "Kunde/Land"
- "Kunde/Email"
- "Artikelliste/GesamtRabatt"
- "Artikelliste/GesamtBrutto"

## Main Functions

This is a summary of what the key functions do in the script:

- `main()`: The main function that calls all other functions in sequence.
- `dict(sheet,lastRow)`: Replaces country names with their corresponding ISO 3166-1 alpha-2 codes in the "Kunde/Land" column.
- `addTtoDate(sheet,lastRow)`: Adds a 'T' to the end of date strings in the "Bestelldatum/Datum" column.
- `addZtoTime(sheet,lastRow)`: Adds a '+01:00' to the end of time strings in the "Bestelldatum/Zeit" column.
- `concateDateTime(sheet, lastRow)`: Concatenates the date and time strings in "Bestelldatum/Datum" and "Bestelldatum/Zeit" columns.
- `substractRanatt(sheet, lastRow)`: Subtracts the "Artikelliste/GesamtRabatt" value from the "Artikelliste/GesamtBrutto" value in the corresponding row.
- `createGender(sheet,lastRow)`: Creates a gender column based on the data in the "Kunde/Anrede" column.
- `createCurrency(sheet,lastRow)`: Adds a new "currency" column to the data.
- `deleteAmazon(sheet,lastRow)`: Deletes rows where the "Kunde/Email" column includes "marketplace.amazon".
- `deleteZeroValue(sheet)`: Deletes rows where the "Artikelliste/GesamtBrutto" column includes a value of "0.00" or "0".
- `deleteColumns(sheet)`, `deleteColumnsZeit(sheet)`: Deletes any columns not specified in the `required` array.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
