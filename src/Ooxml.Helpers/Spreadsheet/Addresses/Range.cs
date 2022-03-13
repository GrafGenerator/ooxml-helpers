namespace Ooxml.Helpers.Spreadsheet.Addresses;

public class Range
{
    public string List { get; }
    public Address UpperLeft { get; private set; }
    public Address BottomRight { get; private set; }

    public int ColumnCount { get; private set; }
    public int RowCount { get; private set; }

    private Range(string address)
    {
        var processedAddress = address;
        var nameDelimiterPosition = processedAddress.IndexOf('!');
        if (nameDelimiterPosition == -1)
        {
            List = "";
        }
        else
        {
            List = processedAddress.Substring(0, nameDelimiterPosition);
            processedAddress = processedAddress.Substring(nameDelimiterPosition + 1);
        }

        if (string.IsNullOrEmpty(processedAddress))
        {
            throw new ArgumentException("Address contains no valid cell address part.", nameof(address));
        }

        var addressParts = processedAddress.Split(new[] {":"}, StringSplitOptions.None)
            .Select(x => x.Trim())
            .ToArray();

        if (addressParts.Length > 2)
        {
            throw new ArgumentException("Invalid cell address.", nameof(address));
        }

        var addresses = addressParts.Select(x => new Address(x)).ToArray();
        UpperLeft = addresses[0];
        BottomRight = addresses.Length > 1 ? addresses[1] : addresses[0];

        Normalize();
    }

    public void Reallocate(ReallocateDirection direction, int count)
    {
        var growthPoint = direction switch
        {
            ReallocateDirection.Up => UpperLeft.Row.NumericPosition,
            ReallocateDirection.Left => UpperLeft.Column.NumericPosition,
            ReallocateDirection.Down => BottomRight.Row.NumericPosition,
            ReallocateDirection.Right => BottomRight.Column.NumericPosition,
            _ => throw new ArgumentOutOfRangeException(nameof(direction), direction, "Unknown direction.")
        };

        var adjustedCount = direction == ReallocateDirection.Up || direction == ReallocateDirection.Left
            ? -count
            : count;
            
        if (growthPoint + adjustedCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(count), count, "Reallocated range goes outside of sheet (coordinate value < 1).");
        }

        switch (direction)
        {
            case ReallocateDirection.Up:
                UpperLeft = new Address(UpperLeft.Column.ReferencePosition + UpperLeft.Row.Move(adjustedCount).ReferencePosition);
                break;
            case ReallocateDirection.Left:
                UpperLeft = new Address(UpperLeft.Column.Move(adjustedCount).ReferencePosition + UpperLeft.Row.ReferencePosition);
                break;
            case ReallocateDirection.Down:
                BottomRight = new Address(BottomRight.Column.ReferencePosition + BottomRight.Row.Move(adjustedCount).ReferencePosition);
                break;
            case ReallocateDirection.Right:
                BottomRight = new Address(BottomRight.Column.Move(adjustedCount).ReferencePosition + BottomRight.Row.ReferencePosition);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(direction), direction, "Unknown direction.");
        }
            
        Normalize();
    }

    private void Normalize()
    {
        var addresses = new []{UpperLeft, BottomRight};
        var columns = addresses.Select(x => x.Column).OrderBy(x => x.NumericPosition).ToArray();
        var rows = addresses.Select(x => x.Row).OrderBy(x => x.NumericPosition).ToArray();

        UpperLeft = new Address(columns[0].ReferencePosition + rows[0].ReferencePosition);
        BottomRight = new Address(columns[1].ReferencePosition + rows[1].ReferencePosition);

        ColumnCount = BottomRight.Column.NumericPosition - UpperLeft.Column.NumericPosition + 1;
        RowCount = BottomRight.Row.NumericPosition - UpperLeft.Row.NumericPosition + 1;
    }
        
    public static Range FromString(string address) => new(address);

    #region Equality members

    protected bool Equals(Range other)
    {
        return List == other.List && UpperLeft.Equals(other.UpperLeft) && BottomRight.Equals(other.BottomRight);
    }

    public override bool Equals(object obj)
    {
        if (ReferenceEquals(null, obj)) return false;
        if (ReferenceEquals(this, obj)) return true;
        if (obj.GetType() != this.GetType()) return false;
        return Equals((Range) obj);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = List.GetHashCode();
            hashCode = (hashCode * 397) ^ UpperLeft.GetHashCode();
            hashCode = (hashCode * 397) ^ BottomRight.GetHashCode();
            return hashCode;
        }
    }

    #endregion
}

public enum ReallocateDirection
{
    Up,
    Down,
    Left,
    Right
}