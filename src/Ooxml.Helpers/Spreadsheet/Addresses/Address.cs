namespace Ooxml.Helpers.Spreadsheet.Addresses;

public class Address
{
    private static readonly char[] Digits = {'0', '1', '2', '3', '4', '5', '6', '7', '8', '9'};

    public Address(string address)
    {
        var firstDigitIndex = address.IndexOfAny(Digits);
        if (firstDigitIndex > 0 && address[firstDigitIndex - 1] == '$')
        {
            firstDigitIndex--;
        }
            
        Column = new AddressColumn(address.Substring(0, firstDigitIndex));
        Row = new AddressRow(address.Substring(firstDigitIndex));
    }

    public AddressRow Row { get; }
    public AddressColumn Column { get; }

    #region Equality members

    protected bool Equals(Address other)
    {
        return Row.Equals(other.Row) && Column.Equals(other.Column);
    }

    public override bool Equals(object obj)
    {
        if (ReferenceEquals(null, obj)) return false;
        if (ReferenceEquals(this, obj)) return true;
        if (obj.GetType() != this.GetType()) return false;
        return Equals((Address) obj);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (Row.GetHashCode() * 397) ^ Column.GetHashCode();
        }
    }

    #endregion
}