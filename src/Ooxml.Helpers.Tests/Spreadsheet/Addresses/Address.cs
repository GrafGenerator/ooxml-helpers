using System;
using NUnit.Framework;

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
}