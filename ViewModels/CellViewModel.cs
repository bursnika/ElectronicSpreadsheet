using ElectronicSpreadsheet.Models;
using System.Windows.Input;
using SpreadsheetCell = ElectronicSpreadsheet.Models.Cell;

namespace ElectronicSpreadsheet.ViewModels;


public class CellViewModel : BaseViewModel
{
    private readonly SpreadsheetCell _cell;
    private readonly Action<CellViewModel> _onCellChanged;
    private readonly Func<bool> _getIsExpressionMode;

    public SpreadsheetCell Cell => _cell;

    public string DisplayText => _getIsExpressionMode() ? _cell.Expression : FormatValue(_cell.Value);

    public string EditText
    {
        get => _cell.Expression;
        set
        {
            if (_cell.Expression != value)
            {
                _cell.Expression = value;
                _onCellChanged(this);
            }
        }
    }

    public bool HasError => _cell.HasError;

    public string? ErrorMessage => _cell.ErrorMessage;

    public CellViewModel(SpreadsheetCell cell, Action<CellViewModel> onCellChanged, Func<bool> getIsExpressionMode)
    {
        _cell = cell;
        _onCellChanged = onCellChanged;
        _getIsExpressionMode = getIsExpressionMode;

        _cell.PropertyChanged += (s, e) =>
        {
            if (e.PropertyName == nameof(SpreadsheetCell.Expression))
            {
                OnPropertyChanged(nameof(DisplayText));
                OnPropertyChanged(nameof(EditText));
            }
            else if (e.PropertyName == nameof(SpreadsheetCell.Value))
            {
                OnPropertyChanged(nameof(DisplayText));
            }

            OnPropertyChanged(nameof(HasError));
            OnPropertyChanged(nameof(ErrorMessage));
        };
    }

    public void UpdateExpression(string expression)
    {
        _onCellChanged(this);
    }

    public void RefreshDisplay()
    {
        OnPropertyChanged(nameof(DisplayText));
    }

    private string FormatValue(object? value)
    {
        if (value == null)
            return string.Empty;

        if (value is bool b)
            return b ? "ІСТИНА" : "ХИБНІСТЬ";

        return value.ToString() ?? string.Empty;
    }
}
