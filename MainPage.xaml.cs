using ElectronicSpreadsheet.Models;
using ElectronicSpreadsheet.ViewModels;
using SpreadsheetCell = ElectronicSpreadsheet.Models.Cell;
using Microsoft.Maui.Controls.Shapes;

namespace ElectronicSpreadsheet;

public partial class MainPage : ContentPage
{
    private MainViewModel? _viewModel;

    public MainPage()
    {
        InitializeComponent();

        Loaded += OnPageLoaded;
    }

    private void OnPageLoaded(object? sender, EventArgs e)
    {
        _viewModel = BindingContext as MainViewModel;
        if (_viewModel != null)
        {
            BuildSpreadsheetGrid();
            _viewModel.PropertyChanged += (s, args) =>
            {
                if (args.PropertyName == nameof(MainViewModel.Spreadsheet))
                {
                    BuildSpreadsheetGrid();
                }
            };
        }
    }

    private void BuildSpreadsheetGrid()
    {
        if (_viewModel?.Spreadsheet == null)
            return;

        SpreadsheetGrid.Children.Clear();
        SpreadsheetGrid.RowDefinitions.Clear();
        SpreadsheetGrid.ColumnDefinitions.Clear();

        var spreadsheet = _viewModel.Spreadsheet;

        // Додаємо визначення рядків та стовпців
        SpreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = 50 }); // Заголовок
        for (int i = 0; i < spreadsheet.Rows; i++)
        {
            SpreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = 50 });
        }

        SpreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = 50 }); // Заголовок рядків
        for (int i = 0; i < spreadsheet.Columns; i++)
        {
            SpreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = 120 });
        }

        // Заголовки стовпців (A, B, C, ...) з рожевим дизайном
        for (int col = 0; col < spreadsheet.Columns; col++)
        {
            var border = new Border
            {
                BackgroundColor = Color.FromArgb("#FFB6C1"), // LightPink
                Padding = new Thickness(8),
                Margin = new Thickness(1)
            };
            border.StrokeShape = new RoundRectangle { CornerRadius = 8 };

            var header = new Label
            {
                Text = SpreadsheetCell.GetColumnName(col),
                HorizontalOptions = LayoutOptions.Center,
                VerticalOptions = LayoutOptions.Center,
                FontAttributes = FontAttributes.Bold,
                TextColor = Colors.White,
                FontSize = 14,
                LineBreakMode = LineBreakMode.NoWrap
            };

            border.Content = header;
            SpreadsheetGrid.Add(border, col + 1, 0);
        }

        // Заголовки рядків (1, 2, 3, ...) з рожевим дизайном
        for (int row = 0; row < spreadsheet.Rows; row++)
        {
            var border = new Border
            {
                BackgroundColor = Color.FromArgb("#FFB6C1"), // LightPink
                Padding = new Thickness(8),
                Margin = new Thickness(1)
            };
            border.StrokeShape = new RoundRectangle { CornerRadius = 8 };

            var header = new Label
            {
                Text = (row + 1).ToString(),
                HorizontalOptions = LayoutOptions.Center,
                VerticalOptions = LayoutOptions.Center,
                FontAttributes = FontAttributes.Bold,
                TextColor = Colors.White,
                FontSize = 14
            };

            border.Content = header;
            SpreadsheetGrid.Add(border, 0, row + 1);
        }

        // Клітинки з красивим дизайном
        foreach (var cellVm in _viewModel.CellViewModels)
        {
            var label = new Label
            {
                HorizontalTextAlignment = TextAlignment.Center,
                VerticalTextAlignment = TextAlignment.Center,
                Margin = new Thickness(2),
                Padding = new Thickness(8),
                BackgroundColor = Colors.White,
                TextColor = Colors.Black,
                FontSize = 14
            };

            var entry = new Entry
            {
                HorizontalTextAlignment = TextAlignment.Center,
                VerticalTextAlignment = TextAlignment.Center,
                Margin = new Thickness(2),
                BackgroundColor = Colors.White,
                TextColor = Colors.Black,
                FontSize = 14,
                IsVisible = false
            };

            // Label показує DisplayText (вираз або значення в залежності від режиму)
            label.SetBinding(Label.TextProperty, new Binding(nameof(CellViewModel.DisplayText), source: cellVm));

            // Entry для редагування - завжди показує Expression
            entry.SetBinding(Entry.TextProperty, new Binding(nameof(CellViewModel.EditText), source: cellVm, mode: BindingMode.TwoWay));

            // При кліці на Label - показуємо Entry
            var tapGesture = new TapGestureRecognizer();
            tapGesture.Tapped += (s, e) =>
            {
                label.IsVisible = false;
                entry.IsVisible = true;
                entry.Focus();
            };
            label.GestureRecognizers.Add(tapGesture);

            // При втраті фокусу Entry - повертаємося до Label
            entry.Unfocused += (s, e) =>
            {
                entry.IsVisible = false;
                label.IsVisible = true;
            };

            // Встановлення кольору при помилці (світло-рожевий для помилок)
            cellVm.Cell.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(SpreadsheetCell.HasError))
                {
                    if (cellVm.Cell.HasError)
                    {
                        label.BackgroundColor = Color.FromArgb("#FFE4E1"); // Pastel pink for errors
                        label.TextColor = Color.FromArgb("#FF1493"); // Deep pink text
                        entry.BackgroundColor = Color.FromArgb("#FFE4E1");
                        entry.TextColor = Color.FromArgb("#FF1493");
                    }
                    else
                    {
                        label.BackgroundColor = Colors.White;
                        label.TextColor = Colors.Black;
                        entry.BackgroundColor = Colors.White;
                        entry.TextColor = Colors.Black;
                    }
                }
            };

            SpreadsheetGrid.Add(label, cellVm.Cell.Column + 1, cellVm.Cell.Row + 1);
            SpreadsheetGrid.Add(entry, cellVm.Cell.Column + 1, cellVm.Cell.Row + 1);
        }

        // Кутовий елемент (верхній лівий кут)
        var cornerBorder = new Border
        {
            BackgroundColor = Color.FromArgb("#FFB6C1"), // LightPink
            Margin = new Thickness(1)
        };
        cornerBorder.StrokeShape = new RoundRectangle { CornerRadius = 8 };
        SpreadsheetGrid.Add(cornerBorder, 0, 0);
    }
}
