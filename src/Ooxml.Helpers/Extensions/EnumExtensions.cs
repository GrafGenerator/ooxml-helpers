using Ooxml.Helpers.Spreadsheet.Addresses;

namespace Ooxml.Helpers.Extensions;

public static class EnumExtensions
{
    public static SheetDirection Reverse(this SheetDirection direction) => direction switch
    {
        SheetDirection.Up => SheetDirection.Down,
        SheetDirection.Down => SheetDirection.Up,
        SheetDirection.Left => SheetDirection.Right,
        SheetDirection.Right => SheetDirection.Left,
        _ => throw new ArgumentOutOfRangeException(nameof(direction), direction, null)
    };
}