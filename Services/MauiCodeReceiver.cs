using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Auth.OAuth2.Responses;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace ElectronicSpreadsheet.Services;

/// <summary>
/// OAuth 2.0 code receiver for MAUI that opens the browser using the MAUI Browser API
/// </summary>
public class MauiCodeReceiver : ICodeReceiver, IDisposable
{
    private string? _redirectUri;
    private HttpListener? _httpListener;

    public string RedirectUri
    {
        get
        {
            if (_redirectUri == null)
            {
                // Use a random port to avoid conflicts and permission issues
                _redirectUri = $"http://127.0.0.1:{GetRandomUnusedPort()}/";
            }
            return _redirectUri;
        }
    }

    public async Task<AuthorizationCodeResponseUrl> ReceiveCodeAsync(AuthorizationCodeRequestUrl url,
        CancellationToken taskCancellationToken)
    {
        var authorizationUrl = url.Build().ToString();
        var logMessages = new List<string>();

        try
        {
            logMessages.Add($"[1] Starting OAuth flow");
            logMessages.Add($"[2] Redirect URI: {RedirectUri}");
            logMessages.Add($"[3] Authorization URL: {authorizationUrl}");

            // Start the local HTTP listener
            _httpListener = new HttpListener();

            logMessages.Add($"[3.5] Attempting to bind HttpListener to {RedirectUri}");

            try
            {
                _httpListener.Prefixes.Add(RedirectUri);
                _httpListener.Start();
                logMessages.Add($"[4] ✓ HTTP listener started successfully on {RedirectUri}");
            }
            catch (HttpListenerException hle)
            {
                logMessages.Add($"[4] ✗ HttpListenerException: Code={hle.ErrorCode}, Message={hle.Message}");
                logMessages.Add($"[4] This usually means: port in use, permission denied, or invalid URL");
                throw new Exception($"Failed to start HTTP listener on {RedirectUri}: {hle.Message}", hle);
            }
            catch (Exception bindEx)
            {
                logMessages.Add($"[4] ✗ Exception binding listener: {bindEx.GetType().Name}: {bindEx.Message}");
                throw;
            }

            // Open the browser using MAUI's Browser API instead of Process.Start
            try
            {
                await Browser.OpenAsync(authorizationUrl, BrowserLaunchMode.SystemPreferred);
                logMessages.Add($"[5] Browser opened successfully");
            }
            catch (Exception ex)
            {
                logMessages.Add($"[5] ERROR opening browser: {ex.Message}");
                throw;
            }

            // Wait for the authorization response
            logMessages.Add($"[6] Waiting for callback from Google...");
            logMessages.Add($"[6.1] Expected callback to: {RedirectUri}");

            var context = await _httpListener.GetContextAsync();
            logMessages.Add($"[7] ✓ Callback received!");

            var fullUrl = context.Request.Url?.ToString();
            var queryParams = context.Request.Url?.Query;

            logMessages.Add($"[8] Full callback URL: {fullUrl}");
            logMessages.Add($"[9] Query string: {queryParams}");

            // Parse query params manually to debug
            if (!string.IsNullOrEmpty(queryParams))
            {
                var hasCode = queryParams.Contains("code=");
                var hasError = queryParams.Contains("error=");
                logMessages.Add($"[9.1] Has 'code' parameter: {hasCode}");
                logMessages.Add($"[9.2] Has 'error' parameter: {hasError}");
            }

            // Send a response to the browser
            var response = context.Response;
            string responseString = "<html><head><meta charset='utf-8'></head><body>Authorization complete. You can close this window.</body></html>";
            var buffer = Encoding.UTF8.GetBytes(responseString);
            response.ContentLength64 = buffer.Length;
            var responseOutput = response.OutputStream;
            await responseOutput.WriteAsync(buffer, 0, buffer.Length);
            responseOutput.Close();
            logMessages.Add($"[10] Response sent to browser");

            // Parse and log the result
            var result = new AuthorizationCodeResponseUrl(queryParams ?? string.Empty);
            logMessages.Add($"[11] Authorization code present: {!string.IsNullOrEmpty(result.Code)}");
            logMessages.Add($"[12] Error in response: {result.Error ?? "(none)"}");
            logMessages.Add($"[13] State: {result.State ?? "(none)"}");

            // Stop the HTTP listener now that we've received the callback
            try
            {
                _httpListener?.Stop();
                logMessages.Add($"[14] HTTP listener stopped");
            }
            catch (Exception stopEx)
            {
                logMessages.Add($"[14] Warning: Error stopping listener: {stopEx.Message}");
            }

            // Log everything
            foreach (var msg in logMessages)
            {
                var fullMsg = $"MauiCodeReceiver: {msg}";
                Console.WriteLine(fullMsg);
            }

            return result;
        }
        catch (Exception ex)
        {
            logMessages.Add($"[ERROR] Exception: {ex.GetType().Name}");
            logMessages.Add($"[ERROR] Message: {ex.Message}");
            logMessages.Add($"[ERROR] Stack: {ex.StackTrace}");

            // Clean up listener on error
            try
            {
                _httpListener?.Stop();
            }
            catch
            {
                // Ignore cleanup errors
            }

            foreach (var msg in logMessages)
            {
                Console.WriteLine($"MauiCodeReceiver: {msg}");
            }

            var logsText = string.Join("\n", logMessages);
            throw new Exception($"OAuth FAILED - DEBUG INFO:\n\n{logsText}\n\nOriginal error: {ex.Message}", ex);
        }
    }

    private static int GetRandomUnusedPort()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        try
        {
            listener.Start();
            return ((IPEndPoint)listener.LocalEndpoint).Port;
        }
        finally
        {
            listener.Stop();
        }
    }

    public void Dispose()
    {
        _httpListener?.Stop();
        _httpListener?.Close();
    }
}
