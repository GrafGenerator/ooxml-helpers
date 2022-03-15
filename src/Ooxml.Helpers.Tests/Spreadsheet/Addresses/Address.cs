using System;
using NUnit.Framework;
using Ooxml.Helpers.Spreadsheet.Addresses;

namespace Ooxml.Helpers.Tests.Spreadsheet.Addresses;

[TestFixture]
public class Address: FixtureBase
{
    [TestCase("A1", ExpectedResult = "A,1")]
    [TestCase("AA111", ExpectedResult = "AA,111")]
    [TestCase("$A1", ExpectedResult = "$A,1")]
    [TestCase("A$1", ExpectedResult = "A,$1")]
    [TestCase("$AA111", ExpectedResult = "$AA,111")]
    [TestCase("AA$111", ExpectedResult = "AA,$111")]
    [TestCase("$AA$111", ExpectedResult = "$AA,$111")]
    [TestCase("a1", ExpectedResult = "A,1")]
    public string AddressCases(string reference)
    {
        var address = new Helpers.Spreadsheet.Addresses.Address(reference);
        return $"{address.Column.ReferencePosition},{address.Row.ReferencePosition}";
    }
    
    [TestCase("A")]
    [TestCase("1")]
    [TestCase("$$A1")]
    [TestCase("$A$$1")]
    public void AddressExceptionalCases(string reference)
    {
        Check<Exception>(() => new Helpers.Spreadsheet.Addresses.Address(reference));
    }
    
    [TestCase("C3", SheetDirection.Up, ExpectedResult = "C2")]
    [TestCase("C3", SheetDirection.Down, ExpectedResult = "C4")]
    [TestCase("C3", SheetDirection.Left, ExpectedResult = "B3")]
    [TestCase("C3", SheetDirection.Right, ExpectedResult = "D3")]
    [TestCase("$C$3", SheetDirection.Up, ExpectedResult = "$C$2")]
    [TestCase("$C$3", SheetDirection.Down, ExpectedResult = "$C$4")]
    [TestCase("$C$3", SheetDirection.Left, ExpectedResult = "$B$3")]
    [TestCase("$C$3", SheetDirection.Right, ExpectedResult = "$D$3")]
    public string AddressAdjacentCases(string reference, SheetDirection direction)
    {
        var address = new Helpers.Spreadsheet.Addresses.Address(reference).Adjacent(direction);
        return $"{address.Column.ReferencePosition}{address.Row.ReferencePosition}";
    }
    
    [TestCase("A1", SheetDirection.Up)]
    [TestCase("A1", SheetDirection.Left)]
    [TestCase("$A$1", SheetDirection.Up)]
    [TestCase("$A$1", SheetDirection.Left)]
    public void AddressAdjacentExceptionCases(string reference, SheetDirection direction)
    {
        Check<Exception>(() => new Helpers.Spreadsheet.Addresses.Address(reference).Adjacent(direction));
    }
    
    [TestCase("A1", ExpectedResult = "A1")]
    [TestCase("$A1", ExpectedResult = "$A1")]
    [TestCase("A$1", ExpectedResult = "A$1")]
    [TestCase("$A$1", ExpectedResult = "$A$1")]
    public string AddressReferenceCases(string reference)
    {
        return new Helpers.Spreadsheet.Addresses.Address(reference).Reference;
    }
    
    [TestCase("A1", ExpectedResult = "A1")]
    [TestCase("$A1", ExpectedResult = "A1")]
    [TestCase("A$1", ExpectedResult = "A1")]
    [TestCase("$A$1", ExpectedResult = "A1")]
    public string AddressCleanReferenceCases(string reference)
    {
        return new Helpers.Spreadsheet.Addresses.Address(reference).CleanReference;
    }
}