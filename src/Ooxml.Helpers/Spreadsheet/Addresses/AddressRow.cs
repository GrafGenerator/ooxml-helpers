using System.Diagnostics;

namespace Ooxml.Helpers.Spreadsheet.Addresses;

[DebuggerDisplay("{ReferencePosition} ({NumericPosition})")]
public class AddressRow
{
    public AddressRow(int numericPosition, bool isFixed = false)
    {
        if (numericPosition <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(numericPosition), numericPosition,
                "Row position must be not negative integer number");
        }

        NumericPosition = numericPosition;
        ReferencePosition = (isFixed ? "$" : "") + NumericPosition;
        CleanReferencePosition = NumericPosition.ToString();
        IsFixed = isFixed;
    }
        
    public AddressRow(string referencePosition)
    {
        if (string.IsNullOrEmpty(referencePosition))
        {
            throw new ArgumentException("Column reference string length must be from 1 to 3.",
                nameof(referencePosition));
        }

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

        if (!int.TryParse(position, out var numericValue))
        {
            throw new ArgumentException("Row position must be not negative integer number",
                nameof(referencePosition));
        }

        if (numericValue <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(referencePosition), numericValue,
                "Row position must be not negative integer number");
        }
            
        ReferencePosition = referencePosition;
        CleanReferencePosition = position;
        NumericPosition = numericValue;
    }

    public AddressRow Move(int count) => new(NumericPosition + count, IsFixed);

    /// <summary>
    /// String representation of row position with possible $ mark
    /// </summary>
    public string ReferencePosition { get; }
    
    /// <summary>
    /// String representation of row position without possible $ mark
    /// </summary>
    public string CleanReferencePosition { get; }

    /// <summary>
    ///     Row numeric position (start from 1)
    /// </summary>
    public int NumericPosition { get; }

    public bool IsFixed { get; }

    #region Equality members

    protected bool Equals(AddressRow other)
    {
        return NumericPosition == other.NumericPosition && IsFixed == other.IsFixed;
    }

    public override bool Equals(object obj)
    {
        if (ReferenceEquals(null, obj)) return false;
        if (ReferenceEquals(this, obj)) return true;
        if (obj.GetType() != this.GetType()) return false;
        return Equals((AddressRow) obj);
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