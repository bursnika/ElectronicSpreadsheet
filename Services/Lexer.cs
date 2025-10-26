using ElectronicSpreadsheet.Models;
using System.Numerics;
using System.Text;

namespace ElectronicSpreadsheet.Services;


public class Lexer
{
    private readonly string _input;
    private int _position;

    public Lexer(string input)
    {
        _input = input ?? string.Empty;
        _position = 0;
    }


    public List<Token> Tokenize()
    {
        var tokens = new List<Token>();

        while (_position < _input.Length)
        {
            SkipWhitespace();

            if (_position >= _input.Length)
                break;

            var token = GetNextToken();
            if (token != null)
            {
                tokens.Add(token);
            }
        }

        tokens.Add(new Token(TokenType.End, "", _position));
        return tokens;
    }

    private void SkipWhitespace()
    {
        while (_position < _input.Length && char.IsWhiteSpace(_input[_position]))
        {
            _position++;
        }
    }

    private Token? GetNextToken()
    {
        if (_position >= _input.Length)
            return null;

        var currentChar = _input[_position];
        var startPosition = _position;

        // Числа (підтримка чисел довільної довжини)
        if (char.IsDigit(currentChar))
        {
            return ReadNumber(startPosition);
        }

        // Посилання на клітинки 
        if (char.IsUpper(currentChar))
        {
            return ReadCellReferenceOrFunction(startPosition);
        }

        // Ключові слова (not, max, min, mmax, mmin)
        if (char.IsLower(currentChar))
        {
            return ReadKeyword(startPosition);
        }

        // Оператори та символи
        _position++;
        return currentChar switch
        {
            '+' => new Token(TokenType.Plus, "+", startPosition),
            '-' => new Token(TokenType.Minus, "-", startPosition),
            '*' => new Token(TokenType.Multiply, "*", startPosition),
            '/' => new Token(TokenType.Divide, "/", startPosition),
            '(' => new Token(TokenType.LeftParen, "(", startPosition),
            ')' => new Token(TokenType.RightParen, ")", startPosition),
            ',' => new Token(TokenType.Comma, ",", startPosition),
            '=' => new Token(TokenType.Equal, "=", startPosition),
            '<' => new Token(TokenType.Less, "<", startPosition),
            '>' => new Token(TokenType.Greater, ">", startPosition),
            _ => throw new Exception($"Невідомий символ '{currentChar}' на позиції {startPosition}")
        };
    }

    private Token ReadNumber(int startPosition)
    {
        var sb = new StringBuilder();

        while (_position < _input.Length && char.IsDigit(_input[_position]))
        {
            sb.Append(_input[_position]);
            _position++;
        }

        return new Token(TokenType.Number, sb.ToString(), startPosition);
    }

    private Token ReadCellReferenceOrFunction(int startPosition)
    {
        var sb = new StringBuilder();

        // Читаємо літери (стовпець)
        while (_position < _input.Length && char.IsUpper(_input[_position]))
        {
            sb.Append(_input[_position]);
            _position++;
        }

        // Якщо після літер йдуть цифри - це посилання на клітинку
        if (_position < _input.Length && char.IsDigit(_input[_position]))
        {
            while (_position < _input.Length && char.IsDigit(_input[_position]))
            {
                sb.Append(_input[_position]);
                _position++;
            }
            return new Token(TokenType.CellReference, sb.ToString(), startPosition);
        }

        throw new Exception($"Неправильне посилання на клітинку '{sb}' на позиції {startPosition}");
    }

    private Token ReadKeyword(int startPosition)
    {
        var sb = new StringBuilder();

        while (_position < _input.Length && char.IsLower(_input[_position]))
        {
            sb.Append(_input[_position]);
            _position++;
        }

        var keyword = sb.ToString();

        return keyword switch
        {
            "not" => new Token(TokenType.Not, keyword, startPosition),
            "max" or "min" or "mmax" or "mmin" => new Token(TokenType.Function, keyword, startPosition),
            _ => throw new Exception($"Невідоме ключове слово '{keyword}' на позиції {startPosition}")
        };
    }
}
