using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlForge.Layout;

/// <summary>
/// Measures the usable horizontal space inside a Word <see cref="TableCell"/> in twips and pixels
/// so images can be proportionally scaled down to fit the column width without overflowing.
/// </summary>
internal static class TableCellWidthMeasurer
{
    /// <summary>
    /// 1 twip = 1/1440 inch; Word renders at 96 DPI -> 15 twips per pixel.
    /// </summary>
    public const double TwipsPerPixel = 15.0;

    /// <summary>
    /// Default horizontal margin per side applied by Word when <c>w:tblCellMar</c> is omitted (108 twips).
    /// </summary>
    private const long DefaultMarginTwips = 108;

    public static int? MeasureUsableWidthInPixels(TableCell? cell)
    {
        if (cell is null || MeasureUsableWidthInTwips(cell) is not { } widthInTwips)
        {
            return null;
        }

        var widthInPixels = (int)Math.Floor(widthInTwips / TwipsPerPixel);
        return widthInPixels > 0 ? widthInPixels : null;
    }

    public static long? MeasureUsableWidthInTwips(TableCell cell)
    {
        var width = GetDeclaredWidth(cell) ?? GetGridWidth(cell);

        if (width is null or <= 0)
        {
            return null;
        }

        var usable = width.Value - GetHorizontalMargins(cell);
        return usable > 0 ? usable : null;
    }

    private static long? GetDeclaredWidth(TableCell cell)
        => GetRawDeclaredWidth(cell) is { } width and > 0 ? width : null;

    private static long? GetRawDeclaredWidth(TableCell cell)
    {
        var width = cell.TableCellProperties?.TableCellWidth;

        if (width?.Width?.Value is not { } value)
        {
            return null;
        }

        if (width.Type?.Value == TableWidthUnitValues.Dxa)
        {
            return ParseTwips(value);
        }

        if (width.Type?.Value == TableWidthUnitValues.Pct &&
            ParsePercentage(value) is { } percentage &&
            GetTableWidth(cell) is { } tableWidth)
        {
            return (long)Math.Round(tableWidth * percentage);
        }

        return null;
    }

    private static long? GetTableWidth(TableCell cell)
    {
        var table = cell.Ancestors<Table>().FirstOrDefault();

        if (table is null)
        {
            return null;
        }

        var width = table.GetFirstChild<TableProperties>()?.TableWidth;

        if (width?.Type?.Value == TableWidthUnitValues.Dxa &&
            ParseTwips(width.Width?.Value) is { } value and > 0)
        {
            return value;
        }

        if (width?.Type?.Value == TableWidthUnitValues.Pct &&
            width.Width?.Value is { } pctRaw &&
            ParsePercentage(pctRaw) is { } tblPct &&
            GetAvailablePageWidth(table) is { } pageWidth)
        {
            return (long)Math.Round(pageWidth * tblPct);
        }

        return SumGridColumns(table) ?? GetAvailablePageWidth(table);
    }

    private static long? GetAvailablePageWidth(Table table)
    {
        var body = table.Ancestors<Body>().FirstOrDefault();
        var sectionProps = body?.Elements<SectionProperties>().LastOrDefault();
        var pageWidth = sectionProps?.GetFirstChild<PageSize>()?.Width?.Value;
        if (pageWidth is null or 0)
        {
            return null;
        }

        var pageMargin = sectionProps?.GetFirstChild<PageMargin>();
        var left = (long)(pageMargin?.Left?.Value ?? 0);
        var right = (long)(pageMargin?.Right?.Value ?? 0);
        var available = (long)pageWidth.Value - left - right;
        return available > 0 ? available : null;
    }

    private static long? GetGridWidth(TableCell cell)
    {
        var row = cell.Ancestors<TableRow>().FirstOrDefault();
        var columns = cell
            .Ancestors<Table>()
            .FirstOrDefault()?
            .GetFirstChild<TableGrid>()?
            .Elements<GridColumn>()
            .ToArray();

        if (row is null || columns is null || columns.Length is 0)
        {
            return null;
        }

        var startColumnIndex = 0;

        foreach (var previousCell in row.Elements<TableCell>())
        {
            if (ReferenceEquals(previousCell, cell))
            {
                break;
            }

            startColumnIndex += GetGridSpan(previousCell);
        }

        var spannedColumns = GetGridSpan(cell);
        long total = 0;

        for (var i = startColumnIndex; i < startColumnIndex + spannedColumns && i < columns.Length; i++)
        {
            total += ParseTwips(columns[i].Width?.Value) ?? 0;
        }

        return total > 0 ? total : null;
    }

    private static long? SumGridColumns(Table table)
    {
        var total = table
            .GetFirstChild<TableGrid>()?
            .Elements<GridColumn>()
            .Sum(col => ParseTwips(col.Width?.Value) ?? 0);

        return total > 0 ? total : null;
    }

    private static int GetGridSpan(TableCell cell)
        => Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);

    private static long GetHorizontalMargins(TableCell cell)
    {
        var margins = cell.TableCellProperties?.TableCellMargin;
        var left = ParseTwips(margins?.LeftMargin?.Width?.Value);
        var right = ParseTwips(margins?.RightMargin?.Width?.Value);

        if (left is null || right is null)
        {
            var tableDefault = cell
                .Ancestors<Table>()
                .FirstOrDefault()?
                .GetFirstChild<TableProperties>()?
                .TableCellMarginDefault;

            left ??= ParseTwips(tableDefault?.TableCellLeftMargin?.Width?.Value);
            right ??= ParseTwips(tableDefault?.TableCellRightMargin?.Width?.Value);
        }

        return (left ?? DefaultMarginTwips) + (right ?? DefaultMarginTwips);
    }

    private static long? ParseTwips(string? value)
    {
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var twips) ||
            double.IsNaN(twips) ||
            double.IsInfinity(twips))
        {
            return null;
        }

        return (long)Math.Round(twips);
    }

    private static long? ParseTwips(int? value) => value;

    private static double? ParsePercentage(string value)
    {
        var text = value.AsSpan().Trim();
        var isPercentSuffixed = text.EndsWith("%", StringComparison.Ordinal);

        if (!double.TryParse(
                isPercentSuffixed ? text[..^1] : text,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var number))
        {
            return null;
        }

        var fraction = isPercentSuffixed ? number / 100.0 : number / 5000.0;
        return fraction is > 0.0 and <= 1.0 ? fraction : null;
    }
}
