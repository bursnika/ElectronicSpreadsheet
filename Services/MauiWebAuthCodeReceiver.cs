using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Auth.OAuth2.Responses;
using System.Net;
using System.Net.Sockets;
using Microsoft.Extensions.Logging;

namespace ElectronicSpreadsheet.Services;

/// <summary>
/// Cross-platform OAuth code receiver using platform-specific mechanisms
/// - Mac/iOS: Uses custom URL scheme (works with App Sandbox)
/// - Windows: Uses HttpListener (more reliable on Windows)
/// </summary>
public class MauiWebAuthCodeReceiver : ICodeReceiver, IDisposable
{
    private HttpListener? _httpListener;
    private string? _redirectUri;
    private TaskCompletionSource<AuthorizationCodeResponseUrl>? _tcs;
    private readonly ILogger<MauiWebAuthCodeReceiver> _logger;

    public MauiWebAuthCodeReceiver(ILogger<MauiWebAuthCodeReceiver>? logger = null)
    {
        _logger = logger ?? LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<MauiWebAuthCodeReceiver>();
    }

    public string RedirectUri
    {
        get
        {
            if (_redirectUri == null)
            {
                // Use loopback redirect for Desktop app credentials
                // This works with Google OAuth Desktop app type
                _redirectUri = $"http://127.0.0.1:{GetRandomUnusedPort()}/";
            }
            return _redirectUri;
        }
    }

    public async Task<AuthorizationCodeResponseUrl> ReceiveCodeAsync(
        AuthorizationCodeRequestUrl url,
        CancellationToken taskCancellationToken)
    {
        var authorizationUrl = url.Build().ToString();
        _logger.LogInformation("Starting OAuth flow with redirect URI: {RedirectUri}", RedirectUri);
        _logger.LogDebug("Authorization URL: {AuthUrl}", authorizationUrl);

        // Use HttpListener with loopback redirect (works with Desktop app credentials)
        return await ReceiveCodeViaHttpListenerAsync(authorizationUrl, taskCancellationToken);
    }

    private async Task<AuthorizationCodeResponseUrl> ReceiveCodeViaHttpListenerAsync(
        string authorizationUrl,
        CancellationToken cancellationToken)
    {
        try
        {
            Console.WriteLine("[OAUTH] Using HttpListener for loopback redirect");
            _logger.LogInformation("Using HttpListener (Windows/Android mode)");

            // Start local HTTP server
            _httpListener = new HttpListener();
            _httpListener.Prefixes.Add(RedirectUri);
            _httpListener.Start();
            Console.WriteLine($"[OAUTH] HTTP listener started on: {RedirectUri}");
            _logger.LogInformation("HTTP listener started on: {RedirectUri}", RedirectUri);

            // Open browser
            Console.WriteLine($"[OAUTH] Opening browser: {authorizationUrl.Substring(0, Math.Min(100, authorizationUrl.Length))}...");
            await Browser.OpenAsync(authorizationUrl, BrowserLaunchMode.SystemPreferred);
            Console.WriteLine("[OAUTH] Browser opened for authentication");
            _logger.LogInformation("Browser opened for authentication");

            // Wait for callback - may receive multiple requests, wait for one with query params
            Console.WriteLine("[OAUTH] Waiting for OAuth callback from browser...");
            _logger.LogInformation("Waiting for OAuth callback...");

            HttpListenerContext context;
            string? queryParams = null;
            int attempts = 0;

            // Handle up to 3 requests - browser might make initial request without params
            while (attempts < 3)
            {
                attempts++;
                context = await _httpListener.GetContextAsync();
                queryParams = context.Request.Url?.Query;

                Console.WriteLine($"[OAUTH] Request #{attempts} received");
                Console.WriteLine($"[OAUTH] Full URL: {context.Request.Url}");
                Console.WriteLine($"[OAUTH] Query parameters: {queryParams}");
                _logger.LogInformation("Request #{Attempt}: URL={Url}, Query={Query}", attempts, context.Request.Url, queryParams);

                // If we have query params (either code or error), process this request
                if (!string.IsNullOrEmpty(queryParams) && queryParams.Length > 1)
                {
                    Console.WriteLine("[OAUTH] OAuth callback with parameters received!");
                    _logger.LogInformation("OAuth callback with parameters received");

                    // Send response to browser
                    var response = context.Response;
                    var responseString = "<html><body>Authorization complete. You can close this window.</body></html>";
                    var buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
                    response.ContentLength64 = buffer.Length;
                    await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                    response.OutputStream.Close();
                    break;
                }
                else
                {
                    // Send empty response and wait for next request
                    Console.WriteLine("[OAUTH] Request without parameters, waiting for next request...");
                    var response = context.Response;
                    response.StatusCode = 200;
                    response.Close();
                }
            }

            _logger.LogDebug("Final query parameters: {Query}", queryParams);

            // Stop listener
            _httpListener.Stop();

            // Remove leading ? from query string if present
            var cleanQueryParams = queryParams?.TrimStart('?') ?? string.Empty;
            Console.WriteLine($"[OAUTH] Clean query params: {cleanQueryParams}");

            var result = new AuthorizationCodeResponseUrl(cleanQueryParams);
            Console.WriteLine($"[OAUTH] Parsed code: {result.Code}");
            Console.WriteLine($"[OAUTH] Parsed error: {result.Error}");
            _logger.LogInformation("Successfully obtained authorization code: {HasCode}", !string.IsNullOrEmpty(result.Code));

            return result;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[OAUTH] EXCEPTION: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine($"[OAUTH] Stack trace: {ex.StackTrace}");
            _logger.LogError(ex, "OAuth authentication failed");
            _httpListener?.Stop();
            throw new Exception($"OAuth failed: {ex.Message}", ex);
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
