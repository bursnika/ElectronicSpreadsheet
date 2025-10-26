namespace ElectronicSpreadsheet.Models;

public enum TokenType
{
    Number,          // Числа
    CellReference,   // Посилання на клітинки (A1, B2, etc.)
    Plus,            // +
    Minus,           // -
    Multiply,        // *
    Divide,          // /
    LeftParen,       // (
    RightParen,      // )
    Comma,           // ,
    Function,        // max, min, mmax, mmin
    Equal,           // =
    Less,            // <
    Greater,         // >
    Not,             // not
    End              // Кінець виразу
}
