# Trichine Report Generator 🐖

Made to generate Result Reports for a German Goverment body that is responsible for testing pork samples for the presence of _Trichinella_ parasites to aid in the prevention of the spread of Trichinosis.

These scripts are run using sample data that has been exported from the institute's Laboratory Management System (LIMS) as an Excel Spreadsheet.

!["Sample Data in an Excelsheet"](/assets/repo/excel-example.png "Example of a typical export")

To use, the user simply either has to include a path to the source Excel sheet as the first argument when using in a command line enviroment:
e.g. `python3 Extern-Laborberichtsgenerator.py external_example.xlsx`

**OR,** when using Windows, the user can simply drag-and-drop the Excel Sheet on top of the desired script.

## Features ✨

- Ability to quickly switch between Validators & Client Organisations in UI
- Quickly and easily produces result reports from excel data
- Allows customisation of result reports in simple, user-friendly UI
  - User can select the target address, automatically insert the corrisponding data of the user ect.
- Consistent Formatting and clean, professional design
- Customer and user data stored in easy to edit `json` files for when people come and go from the department or new recipients need to be added ect.

## Usage ⚙️

- All the data included in this repository has been generated for demonstration purposes
- To test the program yourself, simply clone the repo, remove `_example` from any files that have it as a suffix and use the provided test data to validate that the scripts are running without issue