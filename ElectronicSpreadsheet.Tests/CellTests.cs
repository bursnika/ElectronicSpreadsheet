using ElectronicSpreadsheet.Models;

namespace ElectronicSpreadsheet.Tests;

public class CellTests
{
    [Fact]
    public void GetColumnName_SingleLetter_ReturnsCorrectName()
    {
        
        var result = Cell.GetColumnName(0);

        
        Assert.Equal("A", result);
    }

    [Fact]
    public void GetColumnName_DoubleLetter_ReturnsCorrectName()
    {
        
        var resultZ = Cell.GetColumnName(25);
        var resultAA = Cell.GetColumnName(26); 
        var resultAB = Cell.GetColumnName(27); 

        
        Assert.Equal("Z", resultZ);
        Assert.Equal("AA", resultAA);
        Assert.Equal("AB", resultAB);
    }

    [Fact]
    public void GetColumnIndex_SingleLetter_ReturnsCorrectIndex()
    {
        
        var resultA = Cell.GetColumnIndex("A");
        var resultZ = Cell.GetColumnIndex("Z");

        
        Assert.Equal(0, resultA);
        Assert.Equal(25, resultZ);
    }

    [Fact]
    public void GetColumnIndex_DoubleLetter_ReturnsCorrectIndex()
    {
        
        var resultAA = Cell.GetColumnIndex("AA");
        var resultAB = Cell.GetColumnIndex("AB");
        var resultBA = Cell.GetColumnIndex("BA");

        
        Assert.Equal(26, resultAA);
        Assert.Equal(27, resultAB);
        Assert.Equal(52, resultBA);
    }

    [Fact]
    public void CellReference_ReturnsCorrectReference()
    {
    
        var cell = new Cell { Row = 0, Column = 0 };

        var reference = cell.CellReference;

        
        Assert.Equal("A1", reference);
    }

    [Fact]
    public void PropertyChanged_WhenExpressionChanges_RaisesEvent()
    {
        var cell = new Cell();
        var eventRaised = false;
        cell.PropertyChanged += (sender, args) =>
        {
            if (args.PropertyName == nameof(Cell.Expression))
                eventRaised = true;
        };

        cell.Expression = "=5+3";

        
        Assert.True(eventRaised);
    }
}
