using System.Globalization;

namespace SharpCell;

/// <summary>
/// Excel's implicit conversions between scalar values. Every function uses these instead of
/// inventing its own rules. Errors pass through unchanged; arrays and lambdas are not scalars
/// and convert to <c>#VALUE!</c>.
/// </summary>
internal static class Coercion
{
    public static CellValue ToNumber(CellValue value, CultureInfo culture)
    {
        switch (value.Kind)
        {
            case CellValueKind.Number:
            case CellValueKind.Error:
                return value;
            case CellValueKind.Empty:
            case CellValueKind.Missing:
                return CellValue.Number(0);
            case CellValueKind.Boolean:
                return CellValue.Number(value.AsBoolean() ? 1 : 0);
            case CellValueKind.Text:
                return NumberText.TryParse(value.AsText(), culture, out var number)
                    ? CellValue.Number(number)
                    : CellValue.Error(ErrorKind.Value);
            default:
                return CellValue.Error(ErrorKind.Value);
        }
    }

    public static CellValue ToText(CellValue value, CultureInfo culture)
    {
        switch (value.Kind)
        {
            case CellValueKind.Text:
            case CellValueKind.Error:
                return value;
            case CellValueKind.Empty:
            case CellValueKind.Missing:
                return CellValue.Text("");
            case CellValueKind.Boolean:
                return CellValue.Text(value.AsBoolean() ? "TRUE" : "FALSE");
            case CellValueKind.Number:
                return CellValue.Text(NumberText.FormatGeneral(value.AsNumber(), culture));
            default:
                return CellValue.Error(ErrorKind.Value);
        }
    }

    public static CellValue ToBoolean(CellValue value)
    {
        switch (value.Kind)
        {
            case CellValueKind.Boolean:
            case CellValueKind.Error:
                return value;
            case CellValueKind.Empty:
            case CellValueKind.Missing:
                return CellValue.False;
            case CellValueKind.Number:
                return CellValue.Boolean(value.AsNumber() != 0);
            case CellValueKind.Text:
                var text = value.AsText();
                if (text.Equals("TRUE", System.StringComparison.OrdinalIgnoreCase))
                    return CellValue.True;
                if (text.Equals("FALSE", System.StringComparison.OrdinalIgnoreCase))
                    return CellValue.False;
                return CellValue.Error(ErrorKind.Value);
            default:
                return CellValue.Error(ErrorKind.Value);
        }
    }
}
