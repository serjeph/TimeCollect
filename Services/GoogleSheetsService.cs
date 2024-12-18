﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace TimeCollect.Services
{
    /// <summary>
    /// This class provides methods for interacting with the Google Sheets API
    /// </summary>
    public class GoogleSheetsService
    {
        private readonly string _credentialsPath;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoogleSheetsService"/> class.
        /// </summary>
        /// <param name="credentialsPath"></param>
        public GoogleSheetsService(string credentialsPath)
        {
            _credentialsPath = credentialsPath;
        }

        /// <summary>
        /// Creates a new <see cref="SheetsService"/> instance with the neccessary credentials.
        /// </summary>
        /// <returns>A <see cref="SheetsService"/> instance.</returns>
        public async Task<SheetsService> CreateSheetsService()
        {
            UserCredential credential = await GetUserCredential();
            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "TimeCollect",
            });
        }

        /// <summary>
        /// Obtains user credentials for accessing the Google Sheets API
        /// </summary>
        /// <returns>A <see cref="UserCredential"/> instance.</returns>
        private async Task<UserCredential> GetUserCredential()
        {
            using (var stream = new FileStream(_credentialsPath, FileMode.Open, FileAccess.Read))
            {
                string[] scopes = { SheetsService.Scope.SpreadsheetsReadonly };

                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string installDirectory = Path.Combine(appDataPath, "TimeCollect");


                string tokenFilePath = Path.Combine(installDirectory, "token.json");

                return await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(tokenFilePath, true));
            }

        }
    }
}
