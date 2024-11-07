# TimeCollect

TimeCollect is a C# WPF application designed to streamline the process of collecting and consolidating timesheet data from multiple Google Spreadsheets. It automates the extraction, cleaning, and exporting of data, saving you time and effort.

## Features

* Fetches data from multiple Google Spreadsheets with specified sheet names and ranges.
* Cleans the data by removing unnecessary columns.
* Exports the cleaned data to Excel files (.xlsx).
* Inserts the data into a PostgreSQL database.
* Provides a user-friendly interface with progress updates and logging.

## Getting Started

1. Clone the repository: `git clone https://github.com/serjeph/TimeCollect.git`
2. Open the solution in Visual Studio 2022.
3. Install the required NuGet packages:
    * Google.Apis.Sheets.v4
    * ClosedXML
    * Npgsql
    * Newtonsoft.Json
4. Configure your Google API credentials:
    * Create a Google Cloud Platform project.
    * Enable the Google Sheets API.
    * Download the credentials JSON file and place it in the project directory.
5. Update the database connection string in `MainViewModel.cs`.
6. Build and run the application.

## Usage

1. In the "Data" tab, enter the employee names, nicknames, and corresponding Google Spreadsheet IDs.
2. In the "Sheet Names" tab, enter a comma-separated list of sheet names to fetch data from.
3. Click the "Run" button to start the data processing.
4. Monitor the progress bar and log messages for updates.
5. The cleaned data will be exported to Excel files and inserted into the PostgreSQL database.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
