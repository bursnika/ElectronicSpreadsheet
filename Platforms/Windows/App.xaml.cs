using Microsoft.UI.Xaml;
using System.Runtime.InteropServices;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace ElectronicSpreadsheet.WinUI;

/// <summary>
/// Provides application-specific behavior to supplement the default Application class.
/// </summary>
public partial class App : MauiWinUIApplication
{
	[DllImport("kernel32.dll", SetLastError = true)]
	static extern bool AllocConsole();

	[DllImport("kernel32.dll")]
	static extern IntPtr GetConsoleWindow();

	/// <summary>
	/// Initializes the singleton application object.  This is the first line of authored code
	/// executed, and as such is the logical equivalent of main() or WinMain().
	/// </summary>
	public App()
	{
		// Створюємо консоль для відображення логів
		if (GetConsoleWindow() == IntPtr.Zero)
		{
			AllocConsole();
		}

		this.InitializeComponent();
	}

	protected override MauiApp CreateMauiApp() => MauiProgram.CreateMauiApp();
}

