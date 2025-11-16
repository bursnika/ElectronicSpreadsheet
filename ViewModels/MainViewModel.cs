using ElectronicSpreadsheet.Models;
using ElectronicSpreadsheet.Services;
using System.Collections.ObjectModel;
using System.Windows.Input;

namespace ElectronicSpreadsheet.ViewModels;

public class MainViewModel : BaseViewModel
{
    private readonly SpreadsheetService _spreadsheetService;
    private readonly GoogleDriveService _googleDriveService;
    private Spreadsheet _spreadsheet;
    private bool _isExpressionMode;
    private string _statusMessage = "Готово";

    public ObservableCollection<CellViewModel> CellViewModels { get; } = new();

    public Spreadsheet Spreadsheet
    {
        get => _spreadsheet;
        set => SetProperty(ref _spreadsheet, value);
    }

    public bool IsExpressionMode
    {
        get => _isExpressionMode;
        set
        {
            if (SetProperty(ref _isExpressionMode, value))
            {
                UpdateDisplayMode();
                OnPropertyChanged(nameof(ModeText));
            }
        }
    }

    public string ModeText => _isExpressionMode ? "ВИРАЗ" : "ЗНАЧЕННЯ";

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    public ICommand NewSpreadsheetCommand { get; }
    public ICommand AddRowCommand { get; }
    public ICommand RemoveRowCommand { get; }
    public ICommand AddColumnCommand { get; }
    public ICommand RemoveColumnCommand { get; }
    public ICommand SaveToGoogleDriveCommand { get; }
    public ICommand LoadFromGoogleDriveCommand { get; }
    public ICommand ToggleDisplayModeCommand { get; }
    public ICommand ClearAllCommand { get; }

    public MainViewModel()
    {
        _spreadsheetService = new SpreadsheetService();
        _googleDriveService = new GoogleDriveService();
        _spreadsheet = new Spreadsheet();
        _spreadsheet.InitializeCells();

        InitializeCellViewModels();

        // Команди
        NewSpreadsheetCommand = new Command(ExecuteNewSpreadsheet);
        AddRowCommand = new Command(ExecuteAddRow);
        RemoveRowCommand = new Command(ExecuteRemoveRow);
        AddColumnCommand = new Command(ExecuteAddColumn);
        RemoveColumnCommand = new Command(ExecuteRemoveColumn);
        SaveToGoogleDriveCommand = new Command(async () => await ExecuteSaveToGoogleDrive());
        LoadFromGoogleDriveCommand = new Command(async () => await ExecuteLoadFromGoogleDrive());
        ToggleDisplayModeCommand = new Command(ExecuteToggleDisplayMode);
        ClearAllCommand = new Command(ExecuteClearAll);
    }

    private void InitializeCellViewModels()
    {
        CellViewModels.Clear();
        foreach (var cell in _spreadsheet.Cells)
        {
            CellViewModels.Add(new CellViewModel(cell, OnCellChanged, () => _isExpressionMode));
        }
    }


    private void OnCellChanged(CellViewModel cellViewModel)
    {
        try
        {
            _spreadsheetService.UpdateCell(_spreadsheet, cellViewModel.Cell, cellViewModel.Cell.Expression);
            _spreadsheetService.RecalculateAll(_spreadsheet);

            if (cellViewModel.Cell.HasError)
            {
                StatusMessage = $"Помилка в {cellViewModel.Cell.CellReference}: {cellViewModel.Cell.ErrorMessage}";
            }
            else
            {
                StatusMessage = "Готово";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Помилка: {ex.Message}";
        }
    }

    private void ExecuteNewSpreadsheet()
    {
        _spreadsheet = new Spreadsheet();
        _spreadsheet.InitializeCells();
        InitializeCellViewModels();
        OnPropertyChanged(nameof(Spreadsheet));
        StatusMessage = "Створено нову таблицю";
    }

    private void ExecuteAddRow()
    {
        try
        {
            _spreadsheetService.AddRow(_spreadsheet);
            InitializeCellViewModels();
            OnPropertyChanged(nameof(Spreadsheet));
            StatusMessage = $"Додано рядок. Всього рядків: {_spreadsheet.Rows}";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Помилка: {ex.Message}";
        }
    }

    private void ExecuteRemoveRow()
    {
        try
        {
            if (_spreadsheet.Rows > 1)
            {
                _spreadsheetService.RemoveRow(_spreadsheet, _spreadsheet.Rows - 1);
                InitializeCellViewModels();
                OnPropertyChanged(nameof(Spreadsheet));
                StatusMessage = $"Видалено рядок. Всього рядків: {_spreadsheet.Rows}";
            }
            else
            {
                StatusMessage = "Не можна видалити останній рядок";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Помилка: {ex.Message}";
        }
    }

    private void ExecuteAddColumn()
    {
        try
        {
            _spreadsheetService.AddColumn(_spreadsheet);
            InitializeCellViewModels();
            OnPropertyChanged(nameof(Spreadsheet));
            StatusMessage = $"Додано стовпець. Всього стовпців: {_spreadsheet.Columns}";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Помилка: {ex.Message}";
        }
    }

    private void ExecuteRemoveColumn()
    {
        try
        {
            if (_spreadsheet.Columns > 1)
            {
                _spreadsheetService.RemoveColumn(_spreadsheet, _spreadsheet.Columns - 1);
                InitializeCellViewModels();
                OnPropertyChanged(nameof(Spreadsheet));
                StatusMessage = $"Видалено стовпець. Всього стовпців: {_spreadsheet.Columns}";
            }
            else
            {
                StatusMessage = "Не можна видалити останній стовпець";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Помилка: {ex.Message}";
        }
    }

    private async Task ExecuteSaveToGoogleDrive()
    {
        try
        {
            Console.WriteLine("ExecuteSaveToGoogleDrive: Початок збереження");

            // Запитуємо назву таблиці у користувача
            string fileName = await Application.Current.MainPage.DisplayPromptAsync(
                "Зберегти на Google Drive",
                "Введіть назву таблиці:",
                initialValue: _spreadsheet.Name,
                maxLength: 100,
                keyboard: Keyboard.Text);

            if (string.IsNullOrWhiteSpace(fileName))
            {
                Console.WriteLine("ExecuteSaveToGoogleDrive: Користувач скасував збереження");
                StatusMessage = "Збереження скасовано";
                return;
            }

            // Оновлюємо назву таблиці
            _spreadsheet.Name = fileName;
            Console.WriteLine($"ExecuteSaveToGoogleDrive: Назва таблиці: {fileName}");

            // Авторизація якщо потрібно
            if (!_googleDriveService.IsAuthenticated)
            {
                StatusMessage = "Авторизація в Google Drive...";
                Console.WriteLine("ExecuteSaveToGoogleDrive: Початок авторизації");

                var loginSuccess = await _googleDriveService.LoginAsync();
                Console.WriteLine($"ExecuteSaveToGoogleDrive: Результат авторизації: {loginSuccess}");

                if (!loginSuccess)
                {
                    StatusMessage = "Не вдалося авторизуватися в Google Drive";
                    return;
                }
            }

            StatusMessage = "Збереження на Google Drive...";
            Console.WriteLine("ExecuteSaveToGoogleDrive: Початок збереження на Google Drive");

            await _googleDriveService.SaveSpreadsheetAsync(_spreadsheet);

            Console.WriteLine("ExecuteSaveToGoogleDrive: Збереження завершено");
            StatusMessage = $"Збережено: {_spreadsheet.Name}";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ExecuteSaveToGoogleDrive: Помилка - {ex.Message}");
            Console.WriteLine($"ExecuteSaveToGoogleDrive: StackTrace - {ex.StackTrace}");
            StatusMessage = $"Помилка збереження: {ex.Message}";
        }
    }

    private async Task ExecuteLoadFromGoogleDrive()
    {
        try
        {
            Console.WriteLine("ExecuteLoadFromGoogleDrive: Початок завантаження");

            // Авторизація якщо потрібно
            if (!_googleDriveService.IsAuthenticated)
            {
                StatusMessage = "Авторизація в Google Drive...";
                Console.WriteLine("ExecuteLoadFromGoogleDrive: Початок авторизації");

                try
                {
                    var loginSuccess = await _googleDriveService.LoginAsync();
                    Console.WriteLine($"ExecuteLoadFromGoogleDrive: Результат авторизації: {loginSuccess}");

                    if (!loginSuccess)
                    {
                        StatusMessage = "Помилка авторизації";
                        return;
                    }
                }
                catch (Exception authEx)
                {
                    Console.WriteLine($"ExecuteLoadFromGoogleDrive: Exception під час авторизації: {authEx.Message}");
                    Console.WriteLine($"ExecuteLoadFromGoogleDrive: StackTrace: {authEx.StackTrace}");
                    StatusMessage = $"Помилка авторизації: {authEx.Message}";
                    return;
                }
            }

            StatusMessage = "Завантаження списку Google Sheets файлів...";
            var files = await _googleDriveService.GetGoogleSheetsFilesAsync();
            Console.WriteLine($"ExecuteLoadFromGoogleDrive: Знайдено {files.Count} файлів");

            if (files.Count == 0)
            {
                StatusMessage = "Google Sheets файлів не знайдено";
                Console.WriteLine("ExecuteLoadFromGoogleDrive: Немає файлів");

                // Показуємо детальне повідомлення
                var page = Application.Current?.Windows[0]?.Page;
                if (page != null)
                {
                    await page.DisplayAlert(
                        "Немає файлів",
                        "У вашому Google Drive не знайдено файлів Google Sheets.\n\n" +
                        "Створіть файл Google Sheets в браузері і спробуйте знову.",
                        "OK");
                }
                return;
            }

            // Показуємо діалог вибору файлу зі scrollable списком
            Console.WriteLine($"ExecuteLoadFromGoogleDrive: Відкриваємо діалог вибору файлу");
            var mainPage = Application.Current?.Windows[0]?.Page;
            if (mainPage == null)
            {
                StatusMessage = "Помилка: не вдалося отримати доступ до UI";
                Console.WriteLine("ExecuteLoadFromGoogleDrive: mainPage == null");
                return;
            }

            var fileSelectionPage = new Views.FileSelectionPage(files);
            await mainPage.Navigation.PushModalAsync(fileSelectionPage);
            var file = await fileSelectionPage.GetSelectedFileAsync();
            Console.WriteLine($"ExecuteLoadFromGoogleDrive: Вибрано файл: {file?.Name}");

            if (file == null)
            {
                StatusMessage = "Завантаження скасовано";
                Console.WriteLine("ExecuteLoadFromGoogleDrive: Завантаження скасовано");
                return;
            }

            StatusMessage = $"Завантаження {file.Name}...";
            Console.WriteLine($"ExecuteLoadFromGoogleDrive: Завантажуємо файл з ID: {file.Id}");
            var loadedSpreadsheet = await _googleDriveService.LoadSpreadsheetAsync(file.Id);

            if (loadedSpreadsheet == null)
            {
                StatusMessage = "Помилка: не вдалося завантажити таблицю";
                return;
            }

            Console.WriteLine($"ExecuteLoadFromGoogleDrive: Таблиця завантажена");
            _spreadsheet = loadedSpreadsheet;
            _spreadsheetService.RecalculateAll(_spreadsheet);

            InitializeCellViewModels();
            OnPropertyChanged(nameof(Spreadsheet));
            StatusMessage = $"Завантажено з Google Drive: {file.Name}";
            Console.WriteLine("ExecuteLoadFromGoogleDrive: Успішно завантажено");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ExecuteLoadFromGoogleDrive: Помилка - {ex.Message}");
            Console.WriteLine($"ExecuteLoadFromGoogleDrive: StackTrace - {ex.StackTrace}");
            StatusMessage = $"Помилка завантаження: {ex.Message}";
        }
    }

    private void ExecuteToggleDisplayMode()
    {
        IsExpressionMode = !IsExpressionMode;
        StatusMessage = IsExpressionMode ? "Режим: ВИРАЗ" : "Режим: ЗНАЧЕННЯ";
    }

    private void ExecuteClearAll()
    {
        try
        {
            _spreadsheetService.Clear(_spreadsheet);
            StatusMessage = "Таблицю очищено";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Помилка: {ex.Message}";
        }
    }

    private void UpdateDisplayMode()
    {
        foreach (var cellVm in CellViewModels)
        {
            cellVm.RefreshDisplay();
        }
    }

    public void UpdateCellExpression(int row, int column, string expression)
    {
        var cell = _spreadsheet.GetCell(row, column);
        if (cell != null)
        {
            _spreadsheetService.UpdateCell(_spreadsheet, cell, expression);
            _spreadsheetService.RecalculateAll(_spreadsheet);
        }
    }
}
