using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ElectronicSpreadsheet.Models;

public class Spreadsheet : INotifyPropertyChanged
{
    private string _name = "Нова таблиця";
    private int _rows = 10;
    private int _columns = 10;

    public string Name
    {
        get => _name;
        set
        {
            if (_name != value)
            {
                _name = value;
                OnPropertyChanged();
            }
        }
    }

    public int Rows
    {
        get => _rows;
        set
        {
            if (_rows != value && value > 0)
            {
                _rows = value;
                OnPropertyChanged();
            }
        }
    }


    public int Columns
    {
        get => _columns;
        set
        {
            if (_columns != value && value > 0)
            {
                _columns = value;
                OnPropertyChanged();
            }
        }
    }

    public ObservableCollection<Cell> Cells { get; } = new();

    public string? GoogleDriveFileId { get; set; }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }


    public Cell? GetCell(int row, int column)
    {
        return Cells.FirstOrDefault(c => c.Row == row && c.Column == column);
    }


    public Cell? GetCell(string cellReference)
    {
        var match = System.Text.RegularExpressions.Regex.Match(cellReference, @"^([A-Z]+)(\d+)$");
        if (match.Success)
        {
            var column = Cell.GetColumnIndex(match.Groups[1].Value);
            var row = int.Parse(match.Groups[2].Value) - 1;
            return GetCell(row, column);
        }
        return null;
    }

 
    public void InitializeCells()
    {
        Cells.Clear();
        for (int row = 0; row < Rows; row++)
        {
            for (int col = 0; col < Columns; col++)
            {
                Cells.Add(new Cell { Row = row, Column = col });
            }
        }
    }
}
