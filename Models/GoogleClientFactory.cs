using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace DriveAnalytic.Services
{
    public static class GoogleClientFactory
    {
        public static void Logout()
        {
            string dataFolder = Path.Combine(AppContext.BaseDirectory, "data");
            string tokenPath = Path.Combine(dataFolder, "token");

            if (Directory.Exists(tokenPath))
            {
                Directory.Delete(tokenPath, true);
            }
        }
        public static async Task<DriveService> CreateDriveServiceAsync()
        {
            string dataFolder = Path.Combine(AppContext.BaseDirectory, "data");
            string credentialsPath = Path.Combine(dataFolder, "credentials.json");
            string tokenPath = Path.Combine(dataFolder, "token");

            if (!Directory.Exists(dataFolder))
            {
                Directory.CreateDirectory(dataFolder);
            }

            if (!File.Exists(credentialsPath))
            {
                throw new FileNotFoundException(
                    "Ошибка: credentials.json не найден!\n" +
                    "Пожалуйста, положите файл OAuth 2.0 Client ID в папку: " + dataFolder);
            }

            UserCredential credential;
            using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    new[] { DriveService.Scope.DriveReadonly },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(tokenPath, true));
            }

            return new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "DriveAnalytic",
            });

        }
    }
}