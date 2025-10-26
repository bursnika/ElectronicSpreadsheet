namespace ElectronicSpreadsheet.Models;


public class ParseResult
{

    public bool IsSuccess { get; set; }


    public string? ErrorMessage { get; set; }


    public int ErrorPosition { get; set; }


    public object? Value { get; set; }


    public static ParseResult Success(object? value)
    {
        return new ParseResult
        {
            IsSuccess = true,
            Value = value
        };
    }


    public static ParseResult Error(string errorMessage, int errorPosition = 0)
    {
        return new ParseResult
        {
            IsSuccess = false,
            ErrorMessage = errorMessage,
            ErrorPosition = errorPosition
        };
    }
}
