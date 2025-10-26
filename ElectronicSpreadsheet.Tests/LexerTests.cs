using ElectronicSpreadsheet.Services;
using ElectronicSpreadsheet.Models;

namespace ElectronicSpreadsheet.Tests;

public class LexerTests
{
    [Fact]
    public void Tokenize_SimpleExpression_ReturnsCorrectTokens()
    {
        var lexer = new Lexer("5 + 3");

        var tokens = lexer.Tokenize();

        
        Assert.Equal(4, tokens.Count); 
        Assert.Equal(TokenType.Number, tokens[0].Type);
        Assert.Equal("5", tokens[0].Value);
        Assert.Equal(TokenType.Plus, tokens[1].Type);
        Assert.Equal(TokenType.Number, tokens[2].Type);
        Assert.Equal("3", tokens[2].Value);
        Assert.Equal(TokenType.End, tokens[3].Type);
    }

    [Fact]
    public void Tokenize_ExpressionWithCellReference_ReturnsCorrectTokens()
    {
        var lexer = new Lexer("A1 + B2");

        var tokens = lexer.Tokenize();

        
        Assert.Equal(4, tokens.Count); 
        Assert.Equal(TokenType.CellReference, tokens[0].Type);
        Assert.Equal("A1", tokens[0].Value);
        Assert.Equal(TokenType.Plus, tokens[1].Type);
        Assert.Equal(TokenType.CellReference, tokens[2].Type);
        Assert.Equal("B2", tokens[2].Value);
        Assert.Equal(TokenType.End, tokens[3].Type);
    }

    [Fact]
    public void Tokenize_ExpressionWithFunction_ReturnsCorrectTokens()
    {
        var lexer = new Lexer("max(10, 20)");

        var tokens = lexer.Tokenize();

        
        Assert.Equal(7, tokens.Count); 
        Assert.Equal(TokenType.Function, tokens[0].Type);
        Assert.Equal("max", tokens[0].Value);
        Assert.Equal(TokenType.LeftParen, tokens[1].Type);
        Assert.Equal(TokenType.Number, tokens[2].Type);
        Assert.Equal("10", tokens[2].Value);
        Assert.Equal(TokenType.Comma, tokens[3].Type);
        Assert.Equal(TokenType.Number, tokens[4].Type);
        Assert.Equal("20", tokens[4].Value);
        Assert.Equal(TokenType.RightParen, tokens[5].Type);
        Assert.Equal(TokenType.End, tokens[6].Type);
    }
}
