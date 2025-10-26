using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Reflection;
using Google.Apis.Util.Store;

namespace ElectronicSpreadsheet.Services;

public class GoogleDriveService
{
    private DriveService? _driveService;
    private SheetsService? _sheetsService;
    private UserCredential? _credential;
    private readonly string[] _scopes = new[]
    {
        DriveService.Scope.Drive,
        SheetsService.Scope.Spreadsheets
    };

    private static string TokenPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Google.Apis.Auth");


    public bool IsAuthenticated => _credential != null;


    public async Task<bool> LoginAsync()
    {
        return await LoginAsync(false);
    }

    public async Task<bool> LoginAsync(bool forceReauth)
    {
        try
        {
            if (forceReauth && Directory.Exists(TokenPath))
            {
                Console.WriteLine($"LoginAsync: Видалення старих токенів з {TokenPath}");
                try
                {
                    Directory.Delete(TokenPath, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LoginAsync: Помилка видалення токенів: {ex.Message}");
                }
            }

            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "ElectronicSpreadsheet.credentials.json";

            var stream = assembly.GetManifestResourceStream(resourceName);
            Stream credentialStream;

            if (stream == null)
            {
                var credPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "credentials.json");
                if (!File.Exists(credPath))
                {
                    throw new FileNotFoundException("$09; credentials.json =5 7=0945=>");
                }
                credentialStream = File.OpenRead(credPath);
            }
            else
            {
                credentialStream = stream;
            }

            using (credentialStream)
            {
                Console.WriteLine("LoginAsync: Запуск GoogleWebAuthorizationBroker");

                var secrets = GoogleClientSecrets.FromStream(credentialStream).Secrets;
                Console.WriteLine($"LoginAsync: ClientId = {secrets.ClientId}");

                Console.WriteLine($"LoginAsync: Токени зберігаються в: {TokenPath}");
                Console.WriteLine($"LoginAsync: Scopes: {string.Join(", ", _scopes)}");

                var dataStore = new FileDataStore(TokenPath, true);
                _credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    secrets,
                    _scopes,
                    "user",
                    CancellationToken.None,
                    dataStore);

                Console.WriteLine("LoginAsync: Авторизація успішна");
                Console.WriteLine($"LoginAsync: User = {_credential.UserId}");
            }

            _driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = "Electronic Spreadsheet"
            });

            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = "Electronic Spreadsheet"
            });

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"><8;:0 02B>@870FVW: {ex.Message}");
            return false;
        }
    }


    public async Task<List<GoogleDriveFile>> GetGoogleSheetsFilesAsync()
    {
        if (_driveService == null)
        {
            throw new InvalidOperationException("!?>G0B:C ?>B@V1=> 02B>@87C20B8AO");
        }

        var files = new List<GoogleDriveFile>();

        try
        {
            Console.WriteLine("GetGoogleSheetsFilesAsync: Початок запиту до Google Drive");

            var request = _driveService.Files.List();
            request.Q = "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";
            request.Fields = "files(id, name, modifiedTime, size, webViewLink)";
            request.OrderBy = "modifiedTime desc";
            request.PageSize = 100;

            Console.WriteLine($"GetGoogleSheetsFilesAsync: Query = {request.Q}");
            Console.WriteLine($"GetGoogleSheetsFilesAsync: Fields = {request.Fields}");

            var result = await request.ExecuteAsync();

            Console.WriteLine($"GetGoogleSheetsFilesAsync: Отримано файлів: {result.Files?.Count ?? 0}");

            if (result.Files != null)
            {
                foreach (var file in result.Files)
                {
                    Console.WriteLine($"GetGoogleSheetsFilesAsync: Файл - Id={file.Id}, Name={file.Name}");

                    files.Add(new GoogleDriveFile
                    {
                        Id = file.Id,
                        Name = file.Name,
                        ModifiedTime = file.ModifiedTimeDateTimeOffset?.DateTime,
                        Size = file.Size,
                        WebViewLink = file.WebViewLink
                    });
                }
            }
            else
            {
                Console.WriteLine("GetGoogleSheetsFilesAsync: result.Files is null");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Помилка отримання файлів: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            throw;
        }

        Console.WriteLine($"GetGoogleSheetsFilesAsync: Повертаємо {files.Count} файлів");
        return files;
    }

    public async Task<Models.Spreadsheet?> LoadSpreadsheetAsync(string fileId)
    {
        if (_sheetsService == null)
        {
            throw new InvalidOperationException("!?>G0B:C ?>B@V1=> 02B>@87C20B8AO");
        }

        try
        {
            var spreadsheetRequest = _sheetsService.Spreadsheets.Get(fileId);
            var spreadsheetData = await spreadsheetRequest.ExecuteAsync();

            if (spreadsheetData?.Sheets == null || spreadsheetData.Sheets.Count == 0)
            {
                return null;
            }

            var firstSheet = spreadsheetData.Sheets[0];
            var sheetName = firstSheet.Properties.Title;

            var range = $"{sheetName}!A1:Z100"; // '8B0T<> 4> 100 @O4:V2 V 26 AB>2?FV2
            var valuesRequest = _sheetsService.Spreadsheets.Values.Get(fileId, range);
            var valuesData = await valuesRequest.ExecuteAsync();

            var spreadsheet = new Models.Spreadsheet
            {
                Name = spreadsheetData.Properties.Title,
                GoogleDriveFileId = fileId
            };

            int maxRows = valuesData.Values?.Count ?? 0;
            int maxCols = 0;

            if (valuesData.Values != null)
            {
                foreach (var row in valuesData.Values)
                {
                    if (row.Count > maxCols)
                        maxCols = row.Count;
                }
            }

            spreadsheet.Rows = Math.Max(maxRows, 10);
            spreadsheet.Columns = Math.Max(maxCols, 10);

            spreadsheet.InitializeCells();

            if (valuesData.Values != null)
            {
                for (int row = 0; row < valuesData.Values.Count; row++)
                {
                    var rowData = valuesData.Values[row];
                    for (int col = 0; col < rowData.Count; col++)
                    {
                        var cell = spreadsheet.GetCell(row, col);
                        if (cell != null && rowData[col] != null)
                        {
                            cell.Expression = rowData[col].ToString() ?? string.Empty;
                        }
                    }
                }
            }

            return spreadsheet;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"><8;:0 7020=B065==O B01;8FV: {ex.Message}");
            throw;
        }
    }

    public async Task SaveSpreadsheetAsync(Models.Spreadsheet spreadsheet)
    {
        if (_driveService == null)
        {
            throw new InvalidOperationException("!?>G0B:C ?>B@V1=> 02B>@87C20B8AO");
        }

        try
        {
            Console.WriteLine($"SaveSpreadsheetAsync: Початок збереження {spreadsheet.Name}");

            var csv = new System.Text.StringBuilder();
            for (int row = 0; row < spreadsheet.Rows; row++)
            {
                var rowValues = new List<string>();
                for (int col = 0; col < spreadsheet.Columns; col++)
                {
                    var cell = spreadsheet.GetCell(row, col);
                    var value = cell?.Expression ?? string.Empty;
                    if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                    {
                        value = $"\"{value.Replace("\"", "\"\"")}\"";
                    }
                    rowValues.Add(value);
                }
                csv.AppendLine(string.Join(",", rowValues));
            }

            var tempFile = Path.GetTempFileName();
            await File.WriteAllTextAsync(tempFile, csv.ToString());

            try
            {
                var fileMetadata = new Google.Apis.Drive.v3.Data.File
                {
                    Name = spreadsheet.Name,
                    MimeType = "application/vnd.google-apps.spreadsheet"
                };

                using var stream = new FileStream(tempFile, FileMode.Open, FileAccess.Read);
                var request = _driveService.Files.Create(fileMetadata, stream, "text/csv");
                request.Fields = "id";

                var result = await request.UploadAsync();

                if (result.Status != Google.Apis.Upload.UploadStatus.Completed)
                {
                    throw new Exception($"Помилка завантаження: {result.Exception?.Message}");
                }

                spreadsheet.GoogleDriveFileId = request.ResponseBody.Id;
                Console.WriteLine($"SaveSpreadsheetAsync: Файл збережено з ID: {spreadsheet.GoogleDriveFileId}");
            }
            finally
            {
                // Видаляємо тимчасовий файл
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Помилка збереження таблиці: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }

    private async Task UpdateSpreadsheetAsync(Models.Spreadsheet spreadsheet)
    {
        if (_sheetsService == null || string.IsNullOrEmpty(spreadsheet.GoogleDriveFileId))
        {
            throw new InvalidOperationException("Неможливо оновити таблицю");
        }

        try
        {
            Console.WriteLine($"UpdateSpreadsheetAsync: Оновлення файлу {spreadsheet.GoogleDriveFileId}");

            var spreadsheetRequest = _sheetsService.Spreadsheets.Get(spreadsheet.GoogleDriveFileId);
            var spreadsheetData = await spreadsheetRequest.ExecuteAsync();

            if (spreadsheetData?.Sheets == null || spreadsheetData.Sheets.Count == 0)
            {
                throw new Exception("Таблиця не має жодного листа");
            }

            var firstSheet = spreadsheetData.Sheets[0];
            var sheetName = firstSheet.Properties.Title;

            var values = new List<IList<object>>();

            for (int row = 0; row < spreadsheet.Rows; row++)
            {
                var rowData = new List<object>();
                for (int col = 0; col < spreadsheet.Columns; col++)
                {
                    var cell = spreadsheet.GetCell(row, col);
                    rowData.Add(cell?.Expression ?? string.Empty);
                }
                values.Add(rowData);
            }

            var valueRange = new ValueRange
            {
                Values = values
            };

            var range = $"{sheetName}!A1:Z{spreadsheet.Rows}";
            var updateRequest = _sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheet.GoogleDriveFileId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            var updateResponse = await updateRequest.ExecuteAsync();

            Console.WriteLine($"UpdateSpreadsheetAsync: Оновлено {updateResponse.UpdatedCells} клітинок");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Помилка оновлення таблиці: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }

    private async Task UploadFileAsync(string filePath, string mimeType)
    {
        if (_driveService == null)
        {
            throw new InvalidOperationException("Спочатку потрібно авторизуватися");
        }

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("Файл не знайдено", filePath);
        }

        try
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File
            {
                Name = Path.GetFileName(filePath),
            };

            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var request = _driveService.Files.Create(fileMetadata, fileStream, mimeType);
            request.Fields = "id";
            var result = await request.UploadAsync();

            if (result.Status != Google.Apis.Upload.UploadStatus.Completed)
            {
                throw new Exception($"Помилка завантаження файлу: {result.Exception?.Message}");
            }

            Console.WriteLine($"UploadFileAsync: Файл успішно завантажено з ID: {request.ResponseBody?.Id}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Помилка завантаження файлу: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }

    public void Logout()
    {
        _credential = null;
        _driveService?.Dispose();
        _sheetsService?.Dispose();
        _driveService = null;
        _sheetsService = null;
    }
}

public class GoogleDriveFile
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public DateTime? ModifiedTime { get; set; }
    public long? Size { get; set; }
    public string? WebViewLink { get; set; }
}
