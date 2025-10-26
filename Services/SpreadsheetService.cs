using ElectronicSpreadsheet.Models;
using System.Text.Json;
using SpreadsheetCell = ElectronicSpreadsheet.Models.Cell;

namespace ElectronicSpreadsheet.Services;

public class SpreadsheetService
{

    public void UpdateCell(Spreadsheet spreadsheet, SpreadsheetCell cell, string expression)
    {
        cell.Expression = expression;

        // Якщо вираз не починається з = або це тільки =, то це просто текст
        var trimmed = expression.TrimStart();
        if (!trimmed.StartsWith("=") || trimmed.Length == 1)
        {
            cell.HasError = false;
            cell.ErrorMessage = null;
            cell.Value = expression;
            return;
        }

        var parser = new ExpressionParser(spreadsheet, cell.CellReference);
        var syntaxResult = parser.ValidateSyntax(expression);

        if (!syntaxResult.IsSuccess)
        {
            cell.HasError = true;
            cell.ErrorMessage = syntaxResult.ErrorMessage;
            cell.Value = null;
            return;
        }

        var evalParser = new ExpressionParser(spreadsheet, cell.CellReference);
        var evalResult = evalParser.Parse(expression);

        if (evalResult.IsSuccess)
        {
            cell.HasError = false;
            cell.ErrorMessage = null;
            cell.Value = evalResult.Value;
        }
        else
        {
            cell.HasError = true;
            cell.ErrorMessage = evalResult.ErrorMessage;
            cell.Value = null;
        }
    }


    public void RecalculateAll(Spreadsheet spreadsheet)
    {
        foreach (var cell in spreadsheet.Cells)
        {
            if (!string.IsNullOrEmpty(cell.Expression))
            {
                UpdateCell(spreadsheet, cell, cell.Expression);
            }
        }
    }


    public void AddRow(Spreadsheet spreadsheet)
    {
        var newRow = spreadsheet.Rows;
        spreadsheet.Rows++;

        for (int col = 0; col < spreadsheet.Columns; col++)
        {
            spreadsheet.Cells.Add(new SpreadsheetCell { Row = newRow, Column = col });
        }
    }

    public void RemoveRow(Spreadsheet spreadsheet, int rowIndex)
    {
        if (spreadsheet.Rows <= 1)
        {
            throw new Exception("Не можна видалити останній рядок");
        }

        var cellsToRemove = spreadsheet.Cells.Where(c => c.Row == rowIndex).ToList();
        foreach (var cell in cellsToRemove)
        {
            spreadsheet.Cells.Remove(cell);
        }

        // Зсунути рядки
        foreach (var cell in spreadsheet.Cells.Where(c => c.Row > rowIndex))
        {
            cell.Row--;
        }

        spreadsheet.Rows--;
        RecalculateAll(spreadsheet);
    }

    public void AddColumn(Spreadsheet spreadsheet)
    {
        var newColumn = spreadsheet.Columns;
        spreadsheet.Columns++;

        for (int row = 0; row < spreadsheet.Rows; row++)
        {
            spreadsheet.Cells.Add(new SpreadsheetCell { Row = row, Column = newColumn });
        }
    }

    public void RemoveColumn(Spreadsheet spreadsheet, int columnIndex)
    {
        if (spreadsheet.Columns <= 1)
        {
            throw new Exception("Не можна видалити останній стовпець");
        }

        var cellsToRemove = spreadsheet.Cells.Where(c => c.Column == columnIndex).ToList();
        foreach (var cell in cellsToRemove)
        {
            spreadsheet.Cells.Remove(cell);
        }

        // Зсунути стовпці
        foreach (var cell in spreadsheet.Cells.Where(c => c.Column > columnIndex))
        {
            cell.Column--;
        }

        spreadsheet.Columns--;
        RecalculateAll(spreadsheet);
    }

    public string SerializeToJson(Spreadsheet spreadsheet)
    {
        var data = new
        {
            Name = spreadsheet.Name,
            Rows = spreadsheet.Rows,
            Columns = spreadsheet.Columns,
            Cells = spreadsheet.Cells.Select(c => new
            {
                Row = c.Row,
                Column = c.Column,
                Expression = c.Expression
            }).Where(c => !string.IsNullOrEmpty(c.Expression))
        };

        return JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
    }

    public Spreadsheet DeserializeFromJson(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            var spreadsheet = new Spreadsheet
            {
                Name = root.GetProperty("Name").GetString() ?? "Нова таблиця",
                Rows = root.GetProperty("Rows").GetInt32(),
                Columns = root.GetProperty("Columns").GetInt32()
            };

            spreadsheet.InitializeCells();

            if (root.TryGetProperty("Cells", out var cellsElement))
            {
                foreach (var cellElement in cellsElement.EnumerateArray())
                {
                    var row = cellElement.GetProperty("Row").GetInt32();
                    var column = cellElement.GetProperty("Column").GetInt32();
                    var expression = cellElement.GetProperty("Expression").GetString() ?? "";

                    var cell = spreadsheet.GetCell(row, column);
                    if (cell != null)
                    {
                        UpdateCell(spreadsheet, cell, expression);
                    }
                }
            }

            return spreadsheet;
        }
        catch (Exception ex)
        {
            throw new Exception($"Помилка при завантаженні таблиці: {ex.Message}");
        }
    }
    public void Clear(Spreadsheet spreadsheet)
    {
        foreach (var cell in spreadsheet.Cells)
        {
            cell.Expression = string.Empty;
            cell.Value = null;
            cell.HasError = false;
            cell.ErrorMessage = null;
        }
    }
}
