using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Reflection;
using Google.Apis.Util.Store;
using Microsoft.Extensions.Logging;

namespace ElectronicSpreadsheet.Services;

public class GoogleDriveService
{
    private DriveService? _driveService;
    private SheetsService? _sheetsService;
    private UserCredential? _credential;
    private readonly ILogger<GoogleDriveService> _logger;
    private readonly string[] _scopes = new[]
    {
        DriveService.Scope.Drive,
        SheetsService.Scope.Spreadsheets
    };

    public GoogleDriveService(ILogger<GoogleDriveService>? logger = null)
    {
        _logger = logger ?? LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<GoogleDriveService>();
    }

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
            Console.WriteLine($"[AUTH] Starting authentication process. ForceReauth: {forceReauth}");
            _logger.LogInformation("Starting authentication process. ForceReauth: {ForceReauth}", forceReauth);

            if (forceReauth && Directory.Exists(TokenPath))
            {
                _logger.LogInformation("Deleting old tokens from: {TokenPath}", TokenPath);
                try
                {
                    Directory.Delete(TokenPath, true);
                    _logger.LogInformation("Old tokens deleted successfully");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to delete old tokens");
                }
            }

            Stream credentialStream;

            try
            {
                // Try loading from MAUI app package first (for bundled assets)
                credentialStream = await FileSystem.OpenAppPackageFileAsync("credentials.json");
                _logger.LogInformation("Loaded credentials from MAUI app package");
            }
            catch (FileNotFoundException)
            {
                // Fall back to file system (for development)
                var credPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "credentials.json");
                _logger.LogInformation("credentials.json not found in app package, checking file system: {Path}", credPath);

                if (!File.Exists(credPath))
                {
                    _logger.LogError("credentials.json not found at {Path}", credPath);
                    throw new FileNotFoundException("credentials.json file not found. Please add your Google OAuth credentials.", credPath);
                }
                credentialStream = File.OpenRead(credPath);
                _logger.LogInformation("Loaded credentials from file system");
            }

            using (credentialStream)
            {
                var secrets = GoogleClientSecrets.FromStream(credentialStream).Secrets;
                Console.WriteLine($"[AUTH] OAuth ClientId: {secrets.ClientId}");
                Console.WriteLine($"[AUTH] Token storage path: {TokenPath}");
                Console.WriteLine($"[AUTH] Requested scopes: {string.Join(", ", _scopes)}");
                _logger.LogInformation("OAuth ClientId: {ClientId}", secrets.ClientId);
                _logger.LogInformation("Token storage path: {TokenPath}", TokenPath);
                _logger.LogInformation("Requested scopes: {Scopes}", string.Join(", ", _scopes));

                // Use custom code receiver for Mac Catalyst
                var dataStore = new FileDataStore(TokenPath, true);
                using var codeReceiver = new MauiWebAuthCodeReceiver();

                _logger.LogInformation("Using MauiWebAuthCodeReceiver with redirect URI: {RedirectUri}", codeReceiver.RedirectUri);

                var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = secrets,
                    Scopes = _scopes,
                    DataStore = dataStore
                });

                _logger.LogInformation("Starting OAuth authorization flow...");

                var token = await flow.LoadTokenAsync("user", CancellationToken.None);
                if (token == null || (forceReauth && token.IsExpired(flow.Clock)))
                {
                    Console.WriteLine("[AUTH] Token not found or expired. Starting browser authentication...");
                    _logger.LogInformation("Token not found or expired. Starting browser authentication...");

                    var authUrl = flow.CreateAuthorizationCodeRequest(codeReceiver.RedirectUri);
                    authUrl.Scope = string.Join(" ", _scopes);

                    Console.WriteLine("[AUTH] Opening browser for authentication...");
                    var response = await codeReceiver.ReceiveCodeAsync(authUrl, CancellationToken.None);

                    if (!string.IsNullOrEmpty(response.Error))
                    {
                        Console.WriteLine($"[AUTH] ERROR: OAuth error: {response.Error}");
                        _logger.LogError("OAuth error: {Error}", response.Error);
                        throw new Exception($"OAuth authorization failed: {response.Error}");
                    }

                    if (string.IsNullOrEmpty(response.Code))
                    {
                        Console.WriteLine("[AUTH] ERROR: No authorization code received");
                        _logger.LogError("No authorization code received");
                        throw new Exception("Authorization code not received from OAuth provider");
                    }

                    Console.WriteLine("[AUTH] Authorization code received successfully");
                    _logger.LogInformation("Authorization code received successfully");

                    Console.WriteLine("[AUTH] Exchanging code for token...");
                    token = await flow.ExchangeCodeForTokenAsync("user", response.Code, codeReceiver.RedirectUri, CancellationToken.None);
                    Console.WriteLine("[AUTH] Token exchanged successfully");
                    _logger.LogInformation("Token exchanged successfully");
                }
                else
                {
                    Console.WriteLine("[AUTH] Using existing valid token");
                    _logger.LogInformation("Using existing valid token");
                }

                _credential = new UserCredential(flow, "user", token);
                _logger.LogInformation("Authentication successful. User: {UserId}", _credential.UserId);
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

            _logger.LogInformation("Google Drive and Sheets services initialized successfully");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[AUTH] EXCEPTION: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine($"[AUTH] Stack trace: {ex.StackTrace}");
            _logger.LogError(ex, "Authentication failed: {Message}", ex.Message);
            return false;
        }
    }


    public async Task<List<GoogleDriveFile>> GetGoogleSheetsFilesAsync()
    {
        if (_driveService == null)
        {
            _logger.LogError("Attempted to get files but service is not authenticated");
            throw new InvalidOperationException("You must authenticate before accessing Google Drive");
        }

        var files = new List<GoogleDriveFile>();

        try
        {
            _logger.LogInformation("Fetching Google Sheets and Excel files created by user from Drive");

            var request = _driveService.Files.List();
            // Filter for Google Sheets OR Excel files, created by me, not trashed
            request.Q = "(mimeType='application/vnd.google-apps.spreadsheet' or mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') and 'me' in owners and trashed=false";
            request.Fields = "files(id, name, modifiedTime, size, webViewLink)";
            request.OrderBy = "modifiedTime desc";
            request.PageSize = 100;

            _logger.LogDebug("Query: {Query}, Fields: {Fields}", request.Q, request.Fields);

            var result = await request.ExecuteAsync();

            _logger.LogInformation("Retrieved {Count} files from Google Drive", result.Files?.Count ?? 0);

            if (result.Files != null)
            {
                foreach (var file in result.Files)
                {
                    _logger.LogDebug("File found - Id: {Id}, Name: {Name}", file.Id, file.Name);

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
                _logger.LogWarning("Files list is null in API response");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to retrieve Google Sheets files");
            throw;
        }

        _logger.LogInformation("Returning {Count} files", files.Count);
        return files;
    }

    public async Task<Models.Spreadsheet?> LoadSpreadsheetAsync(string fileId)
    {
        if (_sheetsService == null)
        {
            _logger.LogError("Attempted to load spreadsheet but service is not authenticated");
            throw new InvalidOperationException("You must authenticate before accessing Google Sheets");
        }

        try
        {
            _logger.LogInformation("Loading spreadsheet with ID: {FileId}", fileId);

            var spreadsheetRequest = _sheetsService.Spreadsheets.Get(fileId);
            var spreadsheetData = await spreadsheetRequest.ExecuteAsync();

            _logger.LogDebug("Spreadsheet metadata retrieved: {Title}", spreadsheetData.Properties.Title);

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

            _logger.LogInformation("Spreadsheet loaded successfully: {Name}", spreadsheet.Name);
            return spreadsheet;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load spreadsheet with ID: {FileId}", fileId);
            throw;
        }
    }

    public async Task SaveSpreadsheetAsync(Models.Spreadsheet spreadsheet)
    {
        if (_driveService == null)
        {
            _logger.LogError("Attempted to save spreadsheet but service is not authenticated");
            throw new InvalidOperationException("You must authenticate before saving to Google Drive");
        }

        try
        {
            _logger.LogInformation("Starting save operation for spreadsheet: {Name}", spreadsheet.Name);

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
                _logger.LogInformation("Spreadsheet saved successfully with ID: {FileId}", spreadsheet.GoogleDriveFileId);
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                    _logger.LogDebug("Temporary file deleted: {TempFile}", tempFile);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to save spreadsheet: {Name}", spreadsheet.Name);
            throw;
        }
    }

    private async Task UpdateSpreadsheetAsync(Models.Spreadsheet spreadsheet)
    {
        if (_sheetsService == null || string.IsNullOrEmpty(spreadsheet.GoogleDriveFileId))
        {
            _logger.LogError("Cannot update spreadsheet - service not authenticated or no file ID");
            throw new InvalidOperationException("Cannot update spreadsheet");
        }

        try
        {
            _logger.LogInformation("Updating spreadsheet file: {FileId}", spreadsheet.GoogleDriveFileId);

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

            _logger.LogInformation("Spreadsheet updated successfully. Updated {CellCount} cells", updateResponse.UpdatedCells);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to update spreadsheet: {FileId}", spreadsheet.GoogleDriveFileId);
            throw;
        }
    }

    private async Task UploadFileAsync(string filePath, string mimeType)
    {
        if (_driveService == null)
        {
            _logger.LogError("Attempted to upload file but service is not authenticated");
            throw new InvalidOperationException("You must authenticate before uploading files");
        }

        if (!File.Exists(filePath))
        {
            _logger.LogError("File not found: {FilePath}", filePath);
            throw new FileNotFoundException("File not found", filePath);
        }

        try
        {
            _logger.LogInformation("Uploading file: {FilePath}, MIME type: {MimeType}", filePath, mimeType);
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
                _logger.LogError("File upload failed: {Message}", result.Exception?.Message);
                throw new Exception($"File upload failed: {result.Exception?.Message}");
            }

            _logger.LogInformation("File uploaded successfully with ID: {FileId}", request.ResponseBody?.Id);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to upload file: {FilePath}", filePath);
            throw;
        }
    }

    public void Logout()
    {
        _logger.LogInformation("Logging out and disposing Google services");
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
