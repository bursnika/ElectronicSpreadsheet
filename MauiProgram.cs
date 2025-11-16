using Microsoft.Extensions.Logging;

namespace ElectronicSpreadsheet;

public static class MauiProgram
{
	public static MauiApp CreateMauiApp()
	{
		var builder = MauiApp.CreateBuilder();
		builder
			.UseMauiApp<App>()
			.ConfigureFonts(fonts =>
			{
				fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
				fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
			});

#if DEBUG
		builder.Logging.AddDebug();
		builder.Logging.AddConsole();
		builder.Logging.SetMinimumLevel(LogLevel.Debug);

		// Write logs to a file since Mac Catalyst apps don't output to terminal
		try
		{
			// Store logs in app directory for easy access
			var appDir = AppDomain.CurrentDomain.BaseDirectory;
			var logPath = Path.Combine(appDir, "ElectronicSpreadsheet_debug.log");

			var logFile = File.CreateText(logPath);
			logFile.AutoFlush = true;
			Console.SetOut(logFile);
			Console.WriteLine($"=== LOG FILE: {logPath} ===");
			Console.WriteLine($"=== APP STARTED: {DateTime.Now} ===");
			Console.WriteLine($"=== App Directory: {appDir} ===");
		}
		catch (Exception ex)
		{
			System.Diagnostics.Debug.WriteLine($"Failed to create log file: {ex.Message}");
		}
#endif

		return builder.Build();
	}
}
