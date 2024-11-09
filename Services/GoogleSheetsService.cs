using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace TimeCollect.Services
{
    internal class GoogleSheetsService
    {
        private readonly string _credentialsPath;

        public GoogleSheetsService(string credentialsPath)
        {
            _credentialsPath = credentialsPath;
        }

        public async Task<SheetsService> CreateSheetsService()
        {
            var credential = await GetUserCredential();
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "TimeCollect",
            });
        }

        private async Task<UserCredential> GetUserCredential()
        {
            using (var stream = new FileStream(_credentialsPath, FileMode.Open, FileAccess.Read))
            {
                string[] scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                return await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore("token.json", true));
            }
        }
    }
}
