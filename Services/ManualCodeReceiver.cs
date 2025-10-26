using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Auth.OAuth2.Responses;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace ElectronicSpreadsheet.Services;


public class ManualCodeReceiver : ICodeReceiver
{
    private const string LoopbackHost = "http://localhost";

    public string RedirectUri { get; private set; } = string.Empty;

    public async Task<AuthorizationCodeResponseUrl> ReceiveCodeAsync(
        AuthorizationCodeRequestUrl url,
        CancellationToken taskCancellationToken)
    {
        Console.WriteLine("ReceiveCodeAsync: Початок");

        try
        {
            Console.WriteLine("ReceiveCodeAsync: Використовуємо ручний режим");
            return await ReceiveCodeManuallyAsync(url, taskCancellationToken);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ReceiveCodeAsync: Критична помилка - {ex.Message}");
            Console.WriteLine($"ReceiveCodeAsync: StackTrace - {ex.StackTrace}");
            throw;
        }
    }

    private async Task<AuthorizationCodeResponseUrl> ReceiveCodeWithHttpListenerAsync(
        AuthorizationCodeRequestUrl url,
        CancellationToken taskCancellationToken)
    {
        var port = GetRandomUnusedPort();
        RedirectUri = $"{LoopbackHost}:{port}/";

        url.RedirectUri = RedirectUri;
        var authUrl = url.Build().ToString();

        using var listener = new HttpListener();
        listener.Prefixes.Add(RedirectUri);
        listener.Start();

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = authUrl,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Помилка відкриття браузера: {ex.Message}");
            Console.WriteLine($"Відкрийте цей URL в браузері: {authUrl}");
        }

        _ = Task.Run(async () =>
        {
            try
            {
                var mainPage = Application.Current?.Windows[0]?.Page;
                if (mainPage != null)
                {
                    await mainPage.DisplayAlert(
                        "Авторизація Google Drive",
                        "Браузер відкрито для авторизації.\n\n" +
                        "Увійдіть в Google акаунт та дозвольте доступ до Drive та Sheets.\n\n" +
                        "Після авторизації вікно браузера автоматично закриється.",
                        "OK");
                }
            }
            catch { }
        });

        var context = await listener.GetContextAsync();

        var code = context.Request.QueryString.Get("code");
        var error = context.Request.QueryString.Get("error");

        var response = context.Response;
        string responseString;

        if (!string.IsNullOrEmpty(code))
        {
            responseString = "<html><head><title>Успішно!</title></head><body style='font-family: Arial; text-align: center; padding-top: 50px;'>" +
                           "<h1>✓ Авторизація успішна!</h1>" +
                           "<p>Ви можете закрити це вікно та повернутися до застосунку.</p>" +
                           "<script>setTimeout(function(){window.close();}, 2000);</script>" +
                           "</body></html>";
        }
        else
        {
            responseString = "<html><head><title>Помилка</title></head><body style='font-family: Arial; text-align: center; padding-top: 50px;'>" +
                           $"<h1>✗ Помилка авторизації</h1>" +
                           $"<p>{error ?? "Невідома помилка"}</p>" +
                           "<p>Закрийте це вікно та спробуйте ще раз.</p>" +
                           "</body></html>";
        }

        var buffer = Encoding.UTF8.GetBytes(responseString);
        response.ContentLength64 = buffer.Length;
        response.ContentType = "text/html; charset=utf-8";
        await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
        response.OutputStream.Close();

        listener.Stop();

        if (!string.IsNullOrEmpty(error))
        {
            throw new Exception($"Помилка авторизації: {error}");
        }

        if (string.IsNullOrEmpty(code))
        {
            throw new Exception("Код авторизації не отримано");
        }

        return new AuthorizationCodeResponseUrl
        {
            Code = code
        };
    }

    private async Task<AuthorizationCodeResponseUrl> ReceiveCodeManuallyAsync(
        AuthorizationCodeRequestUrl url,
        CancellationToken taskCancellationToken)
    {
        Console.WriteLine("ReceiveCodeManuallyAsync: Використовуємо ручний режим");

        RedirectUri = "http://localhost:8080/";
        url.RedirectUri = RedirectUri;
        var authUrl = url.Build().ToString();

        Console.WriteLine($"ReceiveCodeManuallyAsync: Auth URL = {authUrl}");

        var mainPage = Application.Current?.Windows[0]?.Page;
        if (mainPage == null)
        {
            throw new Exception("Не вдалося отримати доступ до UI");
        }

        await mainPage.DisplayAlert(
            "Авторизація Google Drive",
            "Браузер відкриється для авторизації.\n\n" +
            "1. Увійдіть в Google акаунт\n" +
            "2. Натисніть 'Дозволити'\n" +
            "3. Браузер покаже помилку підключення (це нормально)\n" +
            "4. Скопіюйте КОД з адресного рядка після 'code='\n" +
            "   Приклад: http://localhost:8080/?code=4/ABC123...\n" +
            "   Потрібно: 4/ABC123...\n" +
            "5. Вставте код в наступне вікно",
            "Зрозуміло");

        try
        {
            Console.WriteLine("ReceiveCodeManuallyAsync: Відкриваємо браузер");
            Process.Start(new ProcessStartInfo
            {
                FileName = authUrl,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Помилка відкриття браузера: {ex.Message}");
        }

        // Показуємо prompt для введення коду
        var code = await mainPage.DisplayPromptAsync(
            "Введіть код авторизації",
            "Вставте код з адресного рядка (тільки частину після 'code='):",
            "OK",
            "Скасувати",
            "Код авторизації",
            maxLength: 500,
            keyboard: Keyboard.Default);

        Console.WriteLine($"ReceiveCodeManuallyAsync: Отримано код довжиною {code?.Length ?? 0}");

        if (string.IsNullOrEmpty(code))
        {
            throw new Exception("Авторизацію скасовано");
        }

        var response = new AuthorizationCodeResponseUrl
        {
            Code = code.Trim()
        };

        Console.WriteLine($"ReceiveCodeManuallyAsync: Повертаємо код, RedirectUri = {RedirectUri}");

        return response;
    }


    private static int GetRandomUnusedPort()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }
}
