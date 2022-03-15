using Ooxml.Helpers.Spreadsheet.Addresses;

namespace Ooxml.Helpers.Extensions;

public static class EnumExtensions
{
    public static RangeDirection Reverse(this RangeDirection direction) => direction switch
    {
        RangeDirection.Up => RangeDirection.Down,
        RangeDirection.Down => RangeDirection.Up,
        RangeDirection.Left => RangeDirection.Right,
        RangeDirection.Right => RangeDirection.Left,
        _ => throw new ArgumentOutOfRangeException(nameof(direction), direction, null)
    };
}