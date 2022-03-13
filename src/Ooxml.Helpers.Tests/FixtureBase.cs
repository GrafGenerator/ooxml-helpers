using System;
using FluentAssertions;

namespace Ooxml.Helpers.Tests;

public class FixtureBase
{
    protected static void Check<TException>(Action act) where TException: Exception
    {
        act.Should().Throw<TException>();
    }
}