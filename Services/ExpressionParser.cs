using ElectronicSpreadsheet.Models;

namespace ElectronicSpreadsheet.Services;

public class ExpressionParser
{
    private List<Token> _tokens = new();
    private int _currentTokenIndex;
    private Token CurrentToken => _currentTokenIndex < _tokens.Count ? _tokens[_currentTokenIndex] : _tokens[^1];
    private readonly Spreadsheet? _spreadsheet;
    private readonly HashSet<string> _visitedCells = new(); // Для запобігання циклічним посиланням
    private readonly string? _currentCellRef;

    public ExpressionParser(Spreadsheet? spreadsheet = null, string? currentCellRef = null)
    {
        _spreadsheet = spreadsheet;
        _currentCellRef = currentCellRef;

        if (!string.IsNullOrEmpty(currentCellRef))
        {
            _visitedCells.Add(currentCellRef);
        }
    }
    public ParseResult Parse(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression))
        {
            return ParseResult.Success(null);
        }

        try
        {
            expression = expression.TrimStart();
            if (expression.StartsWith("="))
            {
                expression = expression.Substring(1).TrimStart();
            }

            var lexer = new Lexer(expression);
            _tokens = lexer.Tokenize();
            _currentTokenIndex = 0;

            var result = ParseExpression();

            if (CurrentToken.Type != TokenType.End)
            {
                return ParseResult.Error($"Неочікуваний токен '{CurrentToken.Value}' на позиції {CurrentToken.Position}", CurrentToken.Position);
            }

            return ParseResult.Success(result);
        }
        catch (Exception ex)
        {
            return ParseResult.Error(ex.Message);
        }
    }

    public ParseResult ValidateSyntax(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression))
        {
            return ParseResult.Success(null);
        }

        try
        {
            expression = expression.TrimStart();
            if (expression.StartsWith("="))
            {
                expression = expression.Substring(1).TrimStart();
            }

            var lexer = new Lexer(expression);
            _tokens = lexer.Tokenize();
            _currentTokenIndex = 0;

            ValidateExpressionSyntax();

            if (CurrentToken.Type != TokenType.End)
            {
                return ParseResult.Error($"Неочікуваний токен '{CurrentToken.Value}' на позиції {CurrentToken.Position}", CurrentToken.Position);
            }

            return ParseResult.Success(true);
        }
        catch (Exception ex)
        {
            return ParseResult.Error(ex.Message);
        }
    }

    private object ParseExpression()
    {
        return ParseLogicalExpression();
    }

    private object ParseLogicalExpression()
    {
        var left = ParseComparisonExpression();

        if (CurrentToken.Type == TokenType.Equal ||
            CurrentToken.Type == TokenType.Less ||
            CurrentToken.Type == TokenType.Greater)
        {
            var operatorType = CurrentToken.Type;
            Consume(operatorType);
            var right = ParseComparisonExpression();

            return operatorType switch
            {
                TokenType.Equal => CompareValues(left, right) == 0,
                TokenType.Less => CompareValues(left, right) < 0,
                TokenType.Greater => CompareValues(left, right) > 0,
                _ => throw new Exception("Невідомий оператор порівняння")
            };
        }

        return left;
    }

    private object ParseComparisonExpression()
    {
        var left = ParseTerm();

        while (CurrentToken.Type == TokenType.Plus || CurrentToken.Type == TokenType.Minus)
        {
            var operatorType = CurrentToken.Type;
            Consume(operatorType);
            var right = ParseTerm();

            if (operatorType == TokenType.Plus)
            {
                left = AddValues(left, right);
            }
            else
            {
                left = SubtractValues(left, right);
            }
        }

        return left;
    }

    private object ParseTerm()
    {
        var left = ParseFactor();

        while (CurrentToken.Type == TokenType.Multiply || CurrentToken.Type == TokenType.Divide)
        {
            var operatorType = CurrentToken.Type;
            Consume(operatorType);
            var right = ParseFactor();

            if (operatorType == TokenType.Multiply)
            {
                left = MultiplyValues(left, right);
            }
            else
            {
                left = DivideValues(left, right);
            }
        }

        return left;
    }

    private object ParseFactor()
    {
        // Число
        if (CurrentToken.Type == TokenType.Number)
        {
            var value = decimal.Parse(CurrentToken.Value);
            Consume(TokenType.Number);
            return value;
        }

        // Посилання на клітинку
        if (CurrentToken.Type == TokenType.CellReference)
        {
            var cellRef = CurrentToken.Value;
            Consume(TokenType.CellReference);

            // Перевірка циклічних посилань
            if (_visitedCells.Contains(cellRef))
            {
                var cycle = string.Join(" → ", _visitedCells) + " → " + cellRef;
                throw new Exception($"Циклічне посилання: {cycle}");
            }

            if (_spreadsheet == null)
            {
                throw new Exception("Не вказано таблицю для розв'язання посилань на клітинки");
            }

            var cell = _spreadsheet.GetCell(cellRef);
            if (cell == null)
            {
                throw new Exception($"Клітинка {cellRef} не знайдена");
            }

            // Якщо клітинка має вираз, обчислити його
            if (!string.IsNullOrEmpty(cell.Expression) && cell.Expression.TrimStart().StartsWith("="))
            {
                _visitedCells.Add(cellRef);
                var parser = new ExpressionParser(_spreadsheet);
                parser._visitedCells.UnionWith(_visitedCells);
                var result = parser.Parse(cell.Expression);
                _visitedCells.Remove(cellRef);

                if (!result.IsSuccess)
                {
                    throw new Exception(result.ErrorMessage);
                }

                return result.Value ?? 0;
            }

            return cell.Value ?? 0;
        }

        // Функції
        if (CurrentToken.Type == TokenType.Function)
        {
            return ParseFunction();
        }

        // Оператор not
        if (CurrentToken.Type == TokenType.Not)
        {
            Consume(TokenType.Not);
            var value = ParseFactor();
            return !ToBool(value);
        }

        // Дужки
        if (CurrentToken.Type == TokenType.LeftParen)
        {
            Consume(TokenType.LeftParen);
            var value = ParseExpression();
            Consume(TokenType.RightParen);
            return value;
        }

        throw new Exception($"Неочікуваний токен '{CurrentToken.Value}' на позиції {CurrentToken.Position}");
    }

    private object ParseFunction()
    {
        var functionName = CurrentToken.Value;
        Consume(TokenType.Function);
        Consume(TokenType.LeftParen);

        var arguments = new List<object>();
        arguments.Add(ParseExpression());

        while (CurrentToken.Type == TokenType.Comma)
        {
            Consume(TokenType.Comma);
            arguments.Add(ParseExpression());
        }

        Consume(TokenType.RightParen);

        return functionName switch
        {
            "max" => Max(arguments),
            "min" => Min(arguments),
            "mmax" => MMax(arguments),
            "mmin" => MMin(arguments),
            _ => throw new Exception($"Невідома функція '{functionName}'")
        };
    }

    private void ValidateExpressionSyntax()
    {
        ValidateLogicalExpressionSyntax();
    }

    private void ValidateLogicalExpressionSyntax()
    {
        ValidateComparisonExpressionSyntax();

        if (CurrentToken.Type == TokenType.Equal ||
            CurrentToken.Type == TokenType.Less ||
            CurrentToken.Type == TokenType.Greater)
        {
            var operatorType = CurrentToken.Type;
            Consume(operatorType);
            ValidateComparisonExpressionSyntax();
        }
    }

    private void ValidateComparisonExpressionSyntax()
    {
        ValidateTermSyntax();

        while (CurrentToken.Type == TokenType.Plus || CurrentToken.Type == TokenType.Minus)
        {
            var operatorType = CurrentToken.Type;
            Consume(operatorType);
            ValidateTermSyntax();
        }
    }

    private void ValidateTermSyntax()
    {
        ValidateFactorSyntax();

        while (CurrentToken.Type == TokenType.Multiply || CurrentToken.Type == TokenType.Divide)
        {
            var operatorType = CurrentToken.Type;
            Consume(operatorType);
            ValidateFactorSyntax();
        }
    }

    private void ValidateFactorSyntax()
    {
        if (CurrentToken.Type == TokenType.Number)
        {
            Consume(TokenType.Number);
            return;
        }

        if (CurrentToken.Type == TokenType.CellReference)
        {
            Consume(TokenType.CellReference);
            return;
        }

        if (CurrentToken.Type == TokenType.Function)
        {
            ValidateFunctionSyntax();
            return;
        }

        if (CurrentToken.Type == TokenType.Not)
        {
            Consume(TokenType.Not);
            ValidateFactorSyntax();
            return;
        }

        if (CurrentToken.Type == TokenType.LeftParen)
        {
            Consume(TokenType.LeftParen);
            ValidateExpressionSyntax();
            Consume(TokenType.RightParen);
            return;
        }

        throw new Exception($"Неочікуваний токен '{CurrentToken.Value}' на позиції {CurrentToken.Position}");
    }

    private void ValidateFunctionSyntax()
    {
        Consume(TokenType.Function);
        Consume(TokenType.LeftParen);

        ValidateExpressionSyntax();

        while (CurrentToken.Type == TokenType.Comma)
        {
            Consume(TokenType.Comma);
            ValidateExpressionSyntax();
        }

        Consume(TokenType.RightParen);
    }

    private void Consume(TokenType expectedType)
    {
        if (CurrentToken.Type != expectedType)
        {
            throw new Exception($"Очікувався токен типу {expectedType}, але знайдено {CurrentToken.Type} на позиції {CurrentToken.Position}");
        }
        _currentTokenIndex++;
    }

    // Допоміжні методи для роботи з decimal
    private object AddValues(object left, object right)
    {
        return ToDecimal(left) + ToDecimal(right);
    }

    private object SubtractValues(object left, object right)
    {
        return ToDecimal(left) - ToDecimal(right);
    }

    private object MultiplyValues(object left, object right)
    {
        return ToDecimal(left) * ToDecimal(right);
    }

    private object DivideValues(object left, object right)
    {
        var rightValue = ToDecimal(right);
        if (rightValue == 0)
        {
            throw new Exception("Ділення на нуль");
        }
        return ToDecimal(left) / rightValue;
    }

    private int CompareValues(object left, object right)
    {
        return ToDecimal(left).CompareTo(ToDecimal(right));
    }

    private decimal ToDecimal(object value)
    {
        return value switch
        {
            decimal d => d,
            int i => (decimal)i,
            long l => (decimal)l,
            bool b => b ? 1m : 0m,
            string s when decimal.TryParse(s, out var result) => result,
            _ => 0m
        };
    }

    private bool ToBool(object value)
    {
        return value switch
        {
            bool b => b,
            decimal d => d != 0m,
            int i => i != 0,
            _ => false
        };
    }

    private object Max(List<object> arguments)
    {
        if (arguments.Count != 2)
        {
            throw new Exception("Функція max приймає рівно 2 аргументи");
        }

        var val1 = ToDecimal(arguments[0]);
        var val2 = ToDecimal(arguments[1]);
        return Math.Max(val1, val2);
    }

    private object Min(List<object> arguments)
    {
        if (arguments.Count != 2)
        {
            throw new Exception("Функція min приймає рівно 2 аргументи");
        }

        var val1 = ToDecimal(arguments[0]);
        var val2 = ToDecimal(arguments[1]);
        return Math.Min(val1, val2);
    }

    private object MMax(List<object> arguments)
    {
        if (arguments.Count == 0)
        {
            throw new Exception("Функція mmax потребує принаймні 1 аргумент");
        }

        return arguments.Select(ToDecimal).Max();
    }

    private object MMin(List<object> arguments)
    {
        if (arguments.Count == 0)
        {
            throw new Exception("Функція mmin потребує принаймні 1 аргумент");
        }

        return arguments.Select(ToDecimal).Min();
    }
}
