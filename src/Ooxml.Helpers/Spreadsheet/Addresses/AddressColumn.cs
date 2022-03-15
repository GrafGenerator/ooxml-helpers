namespace Ooxml.Helpers.Spreadsheet.Addresses;

public class AddressColumn
{
    private const string LettersIndexer = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    public AddressColumn(int numericPosition, bool isFixed = false)
    {
        if (numericPosition < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(numericPosition), numericPosition, "Column numeric position must be greater than 0.");
        }

        var result = "";
        var tmp = numericPosition - 1;
        var i = 0;
            
        do
        {
            if(i > 0)
            {
                --tmp;
            }
                
            var reminder = tmp % LettersIndexer.Length;
            tmp /= LettersIndexer.Length;
            result = LettersIndexer[reminder] + result;
            i++;
        } while (tmp > 0);
            
        ReferencePosition = (isFixed ? "$" : "") + result;
        NumericPosition = numericPosition;
        IsFixed = isFixed;
    }
            
    public AddressColumn(string referencePosition)
    {
        var position = referencePosition;
        if (position.StartsWith("$"))
        {
            IsFixed = true;
            position = position.Substring(1);
        }

        if (position.StartsWith("$"))
        {
            throw new ArgumentException("Only one $ sign allowed in the beginning", nameof(referencePosition));
        }

        if (position.Length <= 0 || position.Length > 3)
        {
            throw new ArgumentOutOfRangeException(nameof(referencePosition), position.Length,
                "Column reference length must be from 1 to 3.");
        }

        var prepared = position.ToUpper();
        var result = 0;
        for (var i = 0; i < prepared.Length; i++)
        {
            result *= LettersIndexer.Length;

            var currentDigitIndex = LettersIndexer.IndexOf(prepared[i]);
            if (currentDigitIndex == -1)
            {
                throw new ArgumentException(
                    $"Column reference contains invalid character(s): {prepared[i]}",
                    nameof(referencePosition));
            }

            result += i < prepared.Length - 1 ? currentDigitIndex + 1 : currentDigitIndex;
        }

        ReferencePosition = referencePosition.ToUpper();
        CleanReferencePosition = prepared;
        NumericPosition = result + 1;
    }

    public AddressColumn Move(int count) => new(NumericPosition + count, IsFixed);

    /// <summary>
    /// String representation of column position with possible $ mark
    /// </summary>
    public string ReferencePosition { get; }
    
    /// <summary>
    /// String representation of column position without possible $ mark
    /// </summary>
    public string CleanReferencePosition { get; }

    /// <summary>
    ///     1-based column numeric position
    /// </summary>
    public int NumericPosition { get; }

    public bool IsFixed { get; }

    #region Equality members

    protected bool Equals(AddressColumn other)
    {
        return NumericPosition == other.NumericPosition && IsFixed == other.IsFixed;
    }

    public override bool Equals(object obj)
    {
        if (ReferenceEquals(null, obj)) return false;
        if (ReferenceEquals(this, obj)) return true;
        if (obj.GetType() != this.GetType()) return false;
        return Equals((AddressColumn) obj);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (NumericPosition * 397) ^ IsFixed.GetHashCode();
        }
    }

    #endregion
}