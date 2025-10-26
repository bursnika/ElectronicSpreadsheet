namespace ElectronicSpreadsheet.Models;


public class Token
{
    public TokenType Type { get; set; }
    public string Value { get; set; } = string.Empty;
    public int Position { get; set; }

    public Token(TokenType type, string value, int position)
    {
        Type = type;
        Value = value;
        Position = position;
    }

    public override string ToString()
    {
        return $"{Type}: {Value} at {Position}";
    }
}
