# TimeCollect

TimeCollect is a C# WPF application that simplifies the process of gathering and consolidating timesheet data from multiple Google Spreadsheets. It automates data extraction, cleaning, exporting, and database insertion, increasing efficiency and reducing manual effort.

## Features

* **Data Extraction:** Fetches data from specified Google Spreadsheets and sheet names.
* **Data Cleaning:** Removes unnecessary columns and fills blank cells with "0.00".
* **Excel Export:** Exports cleaned data to individual Excel files for each sheet and a combined file for all sheets.
* **Database Insertion:** Inserts data into a PostgreSQL database.
* **User-Friendly Interface:** Provides a clear and organized UI with tabs for data input, sheet names, settings, and about.
* **Progress Tracking:** Displays progress updates and logs to keep users informed.
* **Error Handling:** Includes robust error handling to prevent crashes and provide informative messages.
* **Credential Management:** Allows users to securely store and manage their Google API credentials.
* **Data Persistence:** Saves and loads employee data to ensure data is not lost between sessions.
* **UI Enhancements:**
    * Loading indicator for visual feedback during data processing.
    * Clear log button for easy log management.
    * Modern styling and formatting for an improved user experience.
    * Input validation to prevent errors.

## Getting Started

1. Clone the repository: `git clone https://github.com/your-username/TimeCollect.git`
2. Open the solution in Visual Studio 2022.
3. Install the required NuGet packages:
    * Google.Apis.Sheets.v4
    * ClosedXML
    * Npgsql
    * Newtonsoft.Json
4. Configure your Google API credentials:
    * Create a Google Cloud Platform project.
    * Enable the Google Sheets API.
    * Download the `credentials.json` file and place it in the project directory.
5. Update the database connection string in `MainViewModel.cs`.
6. Build and run the application.

## Usage

1.  **Data Tab:** Enter employee names, nicknames, and corresponding Google Spreadsheet IDs.
2.  **Sheet Names Tab:** Enter a comma-separated list of sheet names to fetch data from.
3.  **Settings Tab:**
    *   Enter your Google API Client ID and Client Secret.
    *   Click "Save Credentials" to store your credentials securely.
    *   Specify the desired output directory for the Excel files.
4.  **Click the "Run" button to start the data processing.**
5.  Monitor the progress bar and log messages for updates.
6.  The cleaned data will be exported to Excel files in the specified output directory and inserted into the database.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
