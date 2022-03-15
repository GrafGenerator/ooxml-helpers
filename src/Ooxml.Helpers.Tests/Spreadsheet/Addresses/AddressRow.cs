using System;
using NUnit.Framework;

namespace Ooxml.Helpers.Tests.Spreadsheet.Addresses;

[TestFixture]
public class AddressRow: FixtureBase
{
    [TestCase(1, ExpectedResult = "1,nf")]
    [TestCase(100, ExpectedResult = "100,nf")]
    [TestCase(10000000, ExpectedResult = "10000000,nf")]
    [TestCase(1, true, ExpectedResult = "1,f")]
    [TestCase(100, true, ExpectedResult = "100,f")]
    [TestCase(10000000, true, ExpectedResult = "10000000,f")]
    public string RowNumericPositionCases(int number, bool isFixed = false)
    {
        var row = new Helpers.Spreadsheet.Addresses.AddressRow(number, isFixed);
        return $"{row.NumericPosition},{(row.IsFixed ? "f" : "nf")}";
    }
        
    [Test]
    public void RowNumericPositionExceptionalCases()
    {
        Check<ArgumentOutOfRangeException>(() => new Helpers.Spreadsheet.Addresses.AddressRow(0));
        Check<ArgumentOutOfRangeException>(() => new Helpers.Spreadsheet.Addresses.AddressRow(-1));
    }
        
    [TestCase("1", ExpectedResult = 1)]
    [TestCase("100", ExpectedResult = 100)]
    [TestCase("10000000", ExpectedResult = 10000000)]
    public int RowReferencePositionCases(string reference)
    {
        return new Helpers.Spreadsheet.Addresses.AddressRow(reference).NumericPosition;
    }
        
    [TestCase("1", ExpectedResult = "1,nf")]
    [TestCase("100", ExpectedResult = "100,nf")]
    [TestCase("10000000", ExpectedResult = "10000000,nf")]
    [TestCase("$1", ExpectedResult = "1,f")]
    [TestCase("$100", ExpectedResult = "100,f")]
    [TestCase("$10000000", ExpectedResult = "10000000,f")]
    public string RowFixedCases(string reference)
    {
        var row = new Helpers.Spreadsheet.Addresses.AddressRow(reference);
        return $"{row.NumericPosition},{(row.IsFixed ? "f" : "nf")}";
    }
        
    [TestCase("1", 0, ExpectedResult = "1,nf")]
    [TestCase("1", 1, ExpectedResult = "2,nf")]
    [TestCase("1", 100, ExpectedResult = "101,nf")]
    [TestCase("100", -99, ExpectedResult = "1,nf")]
    [TestCase("100", -1, ExpectedResult = "99,nf")]
    [TestCase("100", 0, ExpectedResult = "100,nf")]
    [TestCase("100", 1, ExpectedResult = "101,nf")]
    [TestCase("100", 100, ExpectedResult = "200,nf")]
    [TestCase("$1", 0, ExpectedResult = "1,f")]
    [TestCase("$1", 1, ExpectedResult = "2,f")]
    [TestCase("$1", 100, ExpectedResult = "101,f")]
    [TestCase("$100", -99, ExpectedResult = "1,f")]
    [TestCase("$100", -1, ExpectedResult = "99,f")]
    [TestCase("$100", 0, ExpectedResult = "100,f")]
    [TestCase("$100", 1, ExpectedResult = "101,f")]
    [TestCase("$100", 100, ExpectedResult = "200,f")]
    public string RowMoveCases(string position, int shift)
    {
        var row = new Helpers.Spreadsheet.Addresses.AddressRow(position).Move(shift);
        return $"{row.NumericPosition},{(row.IsFixed ? "f" : "nf")}";
    }

    [TestCase("1", ExpectedResult = "1")]
    [TestCase("$1", ExpectedResult = "1")]
    public string RowCleanReferencePositionCases(string reference)
    {
        return new Helpers.Spreadsheet.Addresses.AddressRow(reference).CleanReferencePosition;
    }
    
    [Test]
    public void RowReferenceExceptionalCases()
    {
        Check<ArgumentException>(() => new Helpers.Spreadsheet.Addresses.AddressRow("abcd"));
        Check<ArgumentException>(() => new Helpers.Spreadsheet.Addresses.AddressRow(""));
        Check<ArgumentOutOfRangeException>(() => new Helpers.Spreadsheet.Addresses.AddressRow("0"));
        Check<ArgumentOutOfRangeException>(() => new Helpers.Spreadsheet.Addresses.AddressRow("-1"));
        Check<ArgumentException>(() => new Helpers.Spreadsheet.Addresses.AddressRow("$$1"));
        Check<ArgumentException>(() => new Helpers.Spreadsheet.Addresses.AddressRow("$"));
        Check<ArgumentOutOfRangeException>(() => new Helpers.Spreadsheet.Addresses.AddressRow("$0"));
        Check<ArgumentOutOfRangeException>(() => new Helpers.Spreadsheet.Addresses.AddressRow("$-1"));
        Check<ArgumentOutOfRangeException>(() => new Helpers.Spreadsheet.Addresses.AddressRow(1).Move(-1));
    }
}