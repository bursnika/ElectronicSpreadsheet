using ElectronicSpreadsheet.Services;
using ElectronicSpreadsheet.Models;

namespace ElectronicSpreadsheet.Tests;

public class ExpressionParserTests
{
    [Fact]
    public void Parse_SimpleAddition_ReturnsCorrectResult()
    {
        var parser = new ExpressionParser();

        var result = parser.Parse("5 + 3");

        
        Assert.True(result.IsSuccess);
        Assert.Equal(8m, result.Value);
    }

    [Fact]
    public void Parse_ComplexExpression_ReturnsCorrectResult()
    {
        var parser = new ExpressionParser();

        var result = parser.Parse("10 + 5 * 2");

        
        Assert.True(result.IsSuccess);
        Assert.Equal(20m, result.Value);
    }

    [Fact]
    public void Parse_DivisionByZero_ReturnsError()
    {
        var parser = new ExpressionParser();

        var result = parser.Parse("10 / 0");

        
        Assert.False(result.IsSuccess);
        Assert.Contains("Ділення на нуль", result.ErrorMessage);
    }

    [Fact]
    public void Parse_MaxFunction_ReturnsCorrectResult()
    {
        var parser = new ExpressionParser();

        var result = parser.Parse("max(15, 25)");

        
        Assert.True(result.IsSuccess);
        Assert.Equal(25m, result.Value);
    }

    [Fact]
    public void Parse_MinFunction_ReturnsCorrectResult()
    {
        var parser = new ExpressionParser();

        var result = parser.Parse("min(15, 25)");

        
        Assert.True(result.IsSuccess);
        Assert.Equal(15m, result.Value);
    }

    [Fact]
    public void Parse_ComparisonEqual_ReturnsCorrectResult()
    {
        var parser = new ExpressionParser();

        var resultTrue = parser.Parse("5 = 5");
        var resultFalse = parser.Parse("5 = 3");

        
        Assert.True(resultTrue.IsSuccess);
        Assert.True((bool)resultTrue.Value);
        Assert.True(resultFalse.IsSuccess);
        Assert.False((bool)resultFalse.Value);
    }
}
