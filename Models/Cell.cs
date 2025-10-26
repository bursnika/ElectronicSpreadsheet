using System.ComponentModel;
using System.Numerics;
using System.Runtime.CompilerServices;

namespace ElectronicSpreadsheet.Models;

public class Cell : INotifyPropertyChanged
{
    private string _expression = string.Empty;
    private object? _value;
    private bool _hasError;
    private string? _errorMessage;

    public int Row { get; set; }

    public int Column { get; set; }
    public string Expression
    {
        get => _expression;
        set
        {
            if (_expression != value)
            {
                _expression = value;
                OnPropertyChanged();
            }
        }
    }

    public object? Value
    {
        get => _value;
        set
        {
            if (_value != value)
            {
                _value = value;
                OnPropertyChanged();
            }
        }
    }

    public bool HasError
    {
        get => _hasError;
        set
        {
            if (_hasError != value)
            {
                _hasError = value;
                OnPropertyChanged();
            }
        }
    }

    public string? ErrorMessage
    {
        get => _errorMessage;
        set
        {
            if (_errorMessage != value)
            {
                _errorMessage = value;
                OnPropertyChanged();
            }
        }
    }


    public string CellReference => $"{GetColumnName(Column)}{Row + 1}";

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public static string GetColumnName(int columnIndex)
    {
        var columnName = string.Empty;
        while (columnIndex >= 0)
        {
            columnName = (char)('A' + (columnIndex % 26)) + columnName;
            columnIndex = columnIndex / 26 - 1;
        }
        return columnName;
    }


    public static int GetColumnIndex(string columnName)
    {
        int columnIndex = 0;
        for (int i = 0; i < columnName.Length; i++)
        {
            columnIndex *= 26;
            columnIndex += columnName[i] - 'A' + 1;
        }
        return columnIndex - 1;
    }
}
