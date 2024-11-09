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

1.  **Prerequisites:**
    *   .NET Framework 4.7.2 or later
    *   PostgreSQL database
    *   Google Cloud Platform project with Google Sheets API enabled

2.  **Clone the repository:** `git clone https://github.com/your-username/TimeCollect.git`

3.  **Open the solution in Visual Studio 2022.**

4.  **Install the required NuGet packages:**
    *   Google.Apis.Sheets.v4
    *   ClosedXML
    *   Npgsql
    *   Newtonsoft.Json
    *   DotNetEnv

5.  **Configure the application:**
    *   Update the database connection string in `DatabaseService.cs`.
    *   Set the desired output directory in `MainViewModel.cs`.
    *   Create a `.env` file in the project's root directory and add your Google API `ClientId` and `ClientSecret`.

6.  **Build and run the application.**

## Usage

1.  **Data Tab:**
    *   Enter employee names, nicknames, and corresponding Google Spreadsheet IDs.
    *   Each employee should have a unique Employee ID.

2.  **Sheet Names Tab:**
    *   Enter a comma-separated list of sheet names to fetch data from.

3.  **Settings Tab:**
    *   Enter your Google API Client ID and Client Secret.
    *   Click "Save Credentials" to store your credentials securely.
    *   Configure the database connection details.
    *   Click "Save Database Settings" to save the database configuration.
    *   Specify the desired output directory for the Excel files.

4.  **Click the "Run" button to start the data processing.**
    *   The application will fetch data from the specified Google Sheets, clean it, and export it to Excel files.
    *   The cleaned data will also be inserted into the PostgreSQL database.

5.  **Monitor the progress bar and log messages for updates.**

6.  **About Tab:**
    *   View information about the application and developer contacts.

## Code Documentation

The codebase is well-documented with XML comments to provide clear explanations of the purpose and usage of classes and methods.

*   **Models:** Contains classes that represent the data (`Employee.cs`, `DatabaseSettings.cs`).
*   **ViewModels:** Contains the `MainViewModel.cs` which handles the application's logic and data.
*   **Services:** Contains classes for interacting with external services (`GoogleSheetsService.cs`, `DatabaseService.cs`).
*   **Helpers:** Contains helper classes for common tasks (`ExcelHelper.cs`).
*   **Converters:** Contains value converters for XAML bindings (`BoolToVisibilityConverter.cs`).
*   **Commands:** Contains the `RelayCommand.cs` for implementing commands.
*   **ValidationRules:** Contains the `RequiredValidationRule.cs` for validating required fields.
*   **Controls:** Contains custom controls (`LoadingIndicator.xaml`).


## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
