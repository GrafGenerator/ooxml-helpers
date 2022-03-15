using System;
using FluentAssertions;
using NUnit.Framework;
using Ooxml.Helpers.Spreadsheet.Addresses;
using Range = Ooxml.Helpers.Spreadsheet.Addresses.Range;

namespace Ooxml.Helpers.Tests.Spreadsheet.Addresses;

[TestFixture]
public class RangeFixture: FixtureBase
{
    [TestCase("A1:A1", ExpectedResult = ",A,1,A,1,1,1")]
    [TestCase("A1:A7", ExpectedResult = ",A,1,A,7,1,7")]
    [TestCase("A7:A1", ExpectedResult = ",A,1,A,7,1,7")]
    [TestCase("A1:F1", ExpectedResult = ",A,1,F,1,6,1")]
    [TestCase("F1:A1", ExpectedResult = ",A,1,F,1,6,1")]
    [TestCase("A1:F7", ExpectedResult = ",A,1,F,7,6,7")]
    [TestCase("F7:A1", ExpectedResult = ",A,1,F,7,6,7")]
    [TestCase("A7:F1", ExpectedResult = ",A,1,F,7,6,7")]
    [TestCase("F1:A7", ExpectedResult = ",A,1,F,7,6,7")]
    [TestCase("A1", ExpectedResult = ",A,1,A,1,1,1")]
    [TestCase("F7", ExpectedResult = ",F,7,F,7,1,1")]
    [TestCase("List1!A1:A1", ExpectedResult = "List1,A,1,A,1,1,1")]
    [TestCase("List1!A1:A7", ExpectedResult = "List1,A,1,A,7,1,7")]
    [TestCase("List1!A7:A1", ExpectedResult = "List1,A,1,A,7,1,7")]
    [TestCase("List1!A1:F1", ExpectedResult = "List1,A,1,F,1,6,1")]
    [TestCase("List1!F1:A1", ExpectedResult = "List1,A,1,F,1,6,1")]
    [TestCase("List1!A1:F7", ExpectedResult = "List1,A,1,F,7,6,7")]
    [TestCase("List1!F7:A1", ExpectedResult = "List1,A,1,F,7,6,7")]
    [TestCase("List1!A7:F1", ExpectedResult = "List1,A,1,F,7,6,7")]
    [TestCase("List1!F1:A7", ExpectedResult = "List1,A,1,F,7,6,7")]
    [TestCase("List1!A1", ExpectedResult = "List1,A,1,A,1,1,1")]
    [TestCase("List1!F7", ExpectedResult = "List1,F,7,F,7,1,1")]
    public string RangeCases(string address)
    {
        var range = Range.FromString(address);
        return
            $"{range.List},{range.UpperLeft.Column.ReferencePosition},{range.UpperLeft.Row.NumericPosition},{range.BottomRight.Column.ReferencePosition},{range.BottomRight.Row.NumericPosition},{range.ColumnCount},{range.RowCount}";
    }
        
    [TestCase("$A1:F7", ExpectedResult = ",$A,1,F,7")]
    [TestCase("A$1:F7", ExpectedResult = ",A,$1,F,7")]
    [TestCase("A1:$F7", ExpectedResult = ",A,1,$F,7")]
    [TestCase("A1:F$7", ExpectedResult = ",A,1,F,$7")]
    [TestCase("$A$1:F7", ExpectedResult = ",$A,$1,F,7")]
    [TestCase("A1:$F$7", ExpectedResult = ",A,1,$F,$7")]
    [TestCase("$A$1:$F$7", ExpectedResult = ",$A,$1,$F,$7")]
    [TestCase("List1!$A1:F7", ExpectedResult = "List1,$A,1,F,7")]
    [TestCase("List1!A$1:F7", ExpectedResult = "List1,A,$1,F,7")]
    [TestCase("List1!A1:$F7", ExpectedResult = "List1,A,1,$F,7")]
    [TestCase("List1!A1:F$7", ExpectedResult = "List1,A,1,F,$7")]
    [TestCase("List1!$A$1:F7", ExpectedResult = "List1,$A,$1,F,7")]
    [TestCase("List1!A1:$F$7", ExpectedResult = "List1,A,1,$F,$7")]
    [TestCase("List1!$A$1:$F$7", ExpectedResult = "List1,$A,$1,$F,$7")]
    public string RangeFixedCases(string address)
    {
        var range = Range.FromString(address);
        return $"{range.List},{range.UpperLeft.Column.ReferencePosition},{range.UpperLeft.Row.ReferencePosition},{range.BottomRight.Column.ReferencePosition},{range.BottomRight.Row.ReferencePosition}";
    }
        
    [TestCase("D4:F7", RangeDirection.Up, -1, ExpectedResult = "D5:F7")]
    [TestCase("D4:F7", RangeDirection.Up, 1, ExpectedResult = "D3:F7")]
    [TestCase("D4:F7", RangeDirection.Left, -1, ExpectedResult = "E4:F7")]
    [TestCase("D4:F7", RangeDirection.Left, 1, ExpectedResult = "C4:F7")]
    [TestCase("D4:F7", RangeDirection.Down, -1, ExpectedResult = "D4:F6")]
    [TestCase("D4:F7", RangeDirection.Down, 1, ExpectedResult = "D4:F8")]
    [TestCase("D4:F7", RangeDirection.Right, -1, ExpectedResult = "D4:E7")]
    [TestCase("D4:F7", RangeDirection.Right, 1, ExpectedResult = "D4:G7")]
    [TestCase("$D$4:$F$7", RangeDirection.Up, -1, ExpectedResult = "$D$5:$F$7")]
    [TestCase("$D$4:$F$7", RangeDirection.Up, 1, ExpectedResult = "$D$3:$F$7")]
    [TestCase("$D$4:$F$7", RangeDirection.Left, -1, ExpectedResult = "$E$4:$F$7")]
    [TestCase("$D$4:$F$7", RangeDirection.Left, 1, ExpectedResult = "$C$4:$F$7")]
    [TestCase("$D$4:$F$7", RangeDirection.Down, -1, ExpectedResult = "$D$4:$F$6")]
    [TestCase("$D$4:$F$7", RangeDirection.Down, 1, ExpectedResult = "$D$4:$F$8")]
    [TestCase("$D$4:$F$7", RangeDirection.Right, -1, ExpectedResult = "$D$4:$E$7")]
    [TestCase("$D$4:$F$7", RangeDirection.Right, 1, ExpectedResult = "$D$4:$G$7")]
    public string RangeReallocateCases(string address, RangeDirection direction, int count)
    {
        var range = Range.FromString(address);
        range.Reallocate(direction, count);
        return $"{range.UpperLeft.Column.ReferencePosition}{range.UpperLeft.Row.ReferencePosition}:{range.BottomRight.Column.ReferencePosition}{range.BottomRight.Row.ReferencePosition}";
    }

    [Test]
    public void RangeReallocateExceptionalCases()
    {
        Check<ArgumentOutOfRangeException>(() =>
        {
            var range = Range.FromString("A1:C3");
            range.Reallocate(RangeDirection.Up, 1);
        });
            
        Check<ArgumentOutOfRangeException>(() =>
        {
            var range = Range.FromString("A1:C3");
            range.Reallocate(RangeDirection.Left, 1);
        });
    }
        
    [TestCase("")]
    [TestCase(" ")]
    [TestCase("List1!")]
    [TestCase("List1! ")]
    [TestCase("A1:B2:C3")]
    [TestCase("List1!A1:B2:C3")]
    [TestCase("List1!$$A1")]
    [TestCase("List1!A$$1")]
    [TestCase("List1!$")]
    public void RangeExceptionalCases(string address)
    {
        Action act = () => Range.FromString(address);
        act.Should().Throw<Exception>();
    }
}