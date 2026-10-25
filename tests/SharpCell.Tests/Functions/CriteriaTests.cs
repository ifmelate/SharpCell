using System.Globalization;
using SharpCell;
using SharpCell.Functions;

namespace SharpCell.Tests.Functions;

public class CriteriaTests
{
    private static readonly string[] Cultures = ["", "en-US", "tr-TR", "az-Latn-AZ", "de-DE", "fr-FR", "sv-SE", "lt-LT", "ja-JP"];

    private static readonly CellValue[] Criteria =
    [
        CellValue.Empty, CellValue.Number(10), CellValue.Number(0), CellValue.Number(1000), CellValue.Boolean(true),
        CellValue.Error(ErrorKind.NA),
        .. new[]
        {
            "", "=", "<>", "<", ">", "<=", ">=", "10", "=10", "<>10", ">5", "<=10", ">=1000", "1,000", "1.000", "$10", "10%",
            "apple", "APPLE", "=apple", "<>apple", "a*", "A?PLE", "a~*b", "<>a~*b", "~~", "~a", "*", "?", "<>*", "?*?",
            "<b", ">=B", ">apple", "<=Z", "<>", "i", "I", "İ", "ı", "<>i", ">i", "<İ", ">=ı", "istanbul", "İSTANBUL",
            "ISTANBUL", "ı*", "İ*", "i?", "ß", "SS", "ss", "straße", "STRASSE", "<ß", ">ss", "STR*E", "stra?e", "é", "é",
            "<é", "a­b", "ab", "<ab", "æ", "AE", "TRUE", "<>TRUE", "#N/A", "<>#N/A", "Region 1", "<>Region 1", "Region 1*",
            "region ?", ">Region 1", "<=Region 2", "Product 2?", "x\0y", "\uD800", "<\uD800", "\u00AD", ">=\u00AD",
        }.Select(CellValue.Text),
    ];

    private static readonly CellValue[] Values =
    [
        CellValue.Empty, CellValue.Number(0), CellValue.Number(10), CellValue.Number(-10), CellValue.Number(1000),
        CellValue.Number(0.1), CellValue.Boolean(true), CellValue.Boolean(false), CellValue.Error(ErrorKind.NA),
        CellValue.Error(ErrorKind.Div0),
        .. new[]
        {
            "", " ", "apple", "APPLE", "Apple", "apricot", "a*b", "axb", "~", "~a", "a~b", "a", "b", "B", "z", "Z", "10", " 10 ",
            "10.0", "1e1", "$10", "10%", "(10)", "1,000", "1.000", "1 000", "i", "I", "İ", "ı", "istanbul", "İSTANBUL",
            "ISTANBUL", "ıstanbul", "Istanbul", "İstanbul", "ß", "SS", "ss", "ẞ", "straße", "STRASSE", "Straße", "é", "é",
            "É", "æ", "AE", "ae", "­", "a­b", "ab", "AB", "TRUE", "true", "#N/A", "Region 1", "Region 10", "Region 2",
            "region 1", "REGION 1", "Region 1 ", "Product 20", "Product 2", "x\0y", "xy", "\uD800",
        }.Select(CellValue.Text),
    ];

    public static TheoryData<string> CultureNames() => new(Cultures);

    [Theory]
    [MemberData(nameof(CultureNames))]
    public void Matches_like_the_plain_comparison_in_every_culture(string name)
    {
        var culture = CultureInfo.GetCultureInfo(name);
        foreach (var criterionValue in Criteria)
        {
            var criterion = Criterion.Parse(criterionValue, culture, DateSystem.Date1900);
            var reference = ReferenceCriterion.Parse(criterionValue, culture, DateSystem.Date1900);
            Assert.Equal(reference.Matches(CellValue.Empty), criterion.MatchesEmpty);

            // Three passes over the same values: the first sees each text once, the others repeat it.
            for (var pass = 0; pass < 3; pass++)
            {
                for (var i = 0; i < Values.Length; i++)
                {
                    var value = Values[pass == 1 ? Values.Length - 1 - i : i];
                    var expected = reference.Matches(value);
                    var actual = criterion.Matches(value);
                    Assert.True(expected == actual,
                        $"culture '{name}', criterion {Show(criterionValue)}, value {Show(value)}: expected {expected}, got {actual}");
                }
            }
        }
    }

    [Theory]
    [InlineData("k1*")]
    [InlineData("k12345")]
    [InlineData("<>k2")]
    [InlineData(">k5")]
    [InlineData("<=K4999")]
    [InlineData("50")]
    public void A_range_of_all_distinct_texts_matches_correctly(string text)
    {
        var culture = CultureInfo.GetCultureInfo("en-US");
        var criterion = Criterion.Parse(CellValue.Text(text), culture, DateSystem.Date1900);
        var reference = ReferenceCriterion.Parse(CellValue.Text(text), culture, DateSystem.Date1900);
        var values = Enumerable.Range(0, 20_000).Select(i => CellValue.Text(i % 3 == 0 ? i.ToString(CultureInfo.InvariantCulture) : "k" + i)).ToArray();

        // Twice: once while the remembered results fill up, once more after they are full.
        for (var pass = 0; pass < 2; pass++)
        {
            foreach (var value in values)
                Assert.Equal(reference.Matches(value), criterion.Matches(value));
        }
    }

    [Theory]
    [InlineData("k1*")]
    [InlineData("<k2000")]
    [InlineData("100")]
    public void A_range_that_repeats_many_texts_and_then_brings_new_ones_matches_correctly(string text)
    {
        var culture = CultureInfo.GetCultureInfo("tr-TR");
        var criterion = Criterion.Parse(CellValue.Text(text), culture, DateSystem.Date1900);
        var reference = ReferenceCriterion.Parse(CellValue.Text(text), culture, DateSystem.Date1900);
        static CellValue Value(int i) => CellValue.Text(i % 7 == 0 ? i.ToString(CultureInfo.InvariantCulture) : "K" + i);
        // 1000 texts three times fill most of the remembered results with plenty of repeats, so
        // the 3000 new texts after them find it full but still worth looking up; then the first again.
        var first = Enumerable.Range(0, 1000).Select(Value).ToArray();
        var values = first.Concat(first).Concat(first).Concat(Enumerable.Range(1000, 3000).Select(Value)).Concat(first).ToArray();

        foreach (var value in values)
            Assert.Equal(reference.Matches(value), criterion.Matches(value));
    }

    [Fact]
    public void Text_criteria_ignore_case_and_use_wildcards_through_COUNTIF()
    {
        var wb = new Workbook();
        var s = wb.AddSheet("S");
        string[] texts = ["Region 1", "REGION 1", "region 1", "Region 10", "Region 2", "a*b", "axb", "~", "10", "Product 2"];
        for (var i = 0; i < texts.Length; i++)
            s[$"A{i + 1}"].Value = texts[i];
        s["A11"].Value = 10;

        Assert.Equal(CellValue.Number(3), wb.Evaluate("=COUNTIF(A1:A11,\"region 1\")"));
        Assert.Equal(CellValue.Number(8), wb.Evaluate("=COUNTIF(A1:A11,\"<>Region 1\")"));
        Assert.Equal(CellValue.Number(4), wb.Evaluate("=COUNTIF(A1:A11,\"REGION 1*\")"));
        Assert.Equal(CellValue.Number(1), wb.Evaluate("=COUNTIF(A1:A11,\"a~*b\")"));
        Assert.Equal(CellValue.Number(1), wb.Evaluate("=COUNTIF(A1:A11,\"~~\")"));
        Assert.Equal(CellValue.Number(2), wb.Evaluate("=COUNTIF(A1:A11,\"=10\")"));
        Assert.Equal(CellValue.Number(10), wb.Evaluate("=COUNTIF(A1:A11,\"<>10\")"));
        Assert.Equal(CellValue.Number(2), wb.Evaluate("=COUNTIF(A1:A11,\">Region 1\")"));
        Assert.Equal(CellValue.Number(5), wb.Evaluate("=COUNTIF(A1:A11,\"<Product 3\")"));
        Assert.Equal(CellValue.Number(9), wb.Evaluate("=COUNTIF(A1:A20,\"\")"));
        Assert.Equal(CellValue.Number(0), wb.Evaluate("=COUNTIF(A1:A20,\">=\")"));
    }

    [Fact]
    public void Case_follows_the_workbook_culture()
    {
        var wb = new Workbook();
        var s = wb.AddSheet("S");
        s["A1"].Value = "I";
        s["A2"].Value = "İ";
        s["A3"].Value = "Istanbul";
        s["A4"].Value = "İstanbul";
        s["B1"].Formula = "=COUNTIF(A1,\"i\")";
        s["C1"].Formula = "=COUNTIF(A2,\"i\")";
        s["D1"].Formula = "=COUNTIF(A3,\"i*\")";
        s["E1"].Formula = "=COUNTIF(A4,\"i*\")";

        wb.Culture = CultureInfo.GetCultureInfo("tr-TR");
        wb.Recalculate();
        Assert.Equal([0.0, 1, 0, 1], Results());

        wb.Culture = CultureInfo.GetCultureInfo("en-US");
        wb.Recalculate();
        Assert.Equal([1.0, 1], [Results()[0], Results()[2]]);

        double[] Results() => [s["B1"].Value.AsNumber(), s["C1"].Value.AsNumber(), s["D1"].Value.AsNumber(), s["E1"].Value.AsNumber()];
    }

    private static string Show(CellValue value) => value.Kind == CellValueKind.Text
        ? "\"" + string.Concat(value.AsText().Select(c => c is < ' ' or > '~' ? $"\\u{(int)c:X4}" : c.ToString())) + "\""
        : value.ToString();

    /// <summary>The criterion matching as it was before matching remembered results: compares every cell afresh.</summary>
    private sealed class ReferenceCriterion
    {
        private enum Operator { Equal, NotEqual, Less, LessEqual, Greater, GreaterEqual }

        private enum OperandKind { Number, Boolean, Error, Text }

        private Operator _op;
        private OperandKind _kind;
        private double _number;
        private bool _boolean;
        private ErrorKind _error;
        private string _text = "";
        private bool _hasWildcards;
        private CultureInfo _culture = CultureInfo.InvariantCulture;

        public static ReferenceCriterion Parse(CellValue criterion, CultureInfo culture, DateSystem dateSystem)
        {
            var c = new ReferenceCriterion { _culture = culture, _op = Operator.Equal };
            switch (criterion.Kind)
            {
                case CellValueKind.Number: c._kind = OperandKind.Number; c._number = criterion.AsNumber(); return c;
                case CellValueKind.Boolean: c._kind = OperandKind.Boolean; c._boolean = criterion.AsBoolean(); return c;
                case CellValueKind.Error: c._kind = OperandKind.Error; c._error = criterion.AsError(); return c;
                case CellValueKind.Text: break;
                default: c._kind = OperandKind.Number; return c;
            }

            var text = criterion.AsText();
            var (op, length) = text switch
            {
                _ when text.StartsWith("<=", StringComparison.Ordinal) => (Operator.LessEqual, 2),
                _ when text.StartsWith(">=", StringComparison.Ordinal) => (Operator.GreaterEqual, 2),
                _ when text.StartsWith("<>", StringComparison.Ordinal) => (Operator.NotEqual, 2),
                _ when text.StartsWith('<') => (Operator.Less, 1),
                _ when text.StartsWith('>') => (Operator.Greater, 1),
                _ when text.StartsWith('=') => (Operator.Equal, 1),
                _ => (Operator.Equal, 0),
            };
            c._op = op;
            var operand = text[length..];
            if (operand.Length > 0)
            {
                var number = Coercion.ToNumber(CellValue.Text(operand), culture, dateSystem);
                if (number.Kind == CellValueKind.Number) { c._kind = OperandKind.Number; c._number = number.AsNumber(); return c; }
                if (operand.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) { c._kind = OperandKind.Boolean; c._boolean = true; return c; }
                if (operand.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) { c._kind = OperandKind.Boolean; return c; }
                if (ErrorKinds.TryParse(operand, out var error)) { c._kind = OperandKind.Error; c._error = error; return c; }
            }

            c._kind = OperandKind.Text;
            c._text = operand;
            c._hasWildcards = operand.AsSpan().IndexOfAny("*?~") >= 0;
            return c;
        }

        public bool Matches(CellValue value)
        {
            if (_op == Operator.NotEqual)
                return !Equal(value);
            if (_op == Operator.Equal)
                return Equal(value);

            int order;
            switch (_kind)
            {
                case OperandKind.Number when value.Kind == CellValueKind.Number:
                    order = value.AsNumber().CompareTo(_number);
                    break;
                case OperandKind.Text when _text.Length == 0:
                    return false;
                case OperandKind.Text when value.Kind == CellValueKind.Text:
                    order = string.Compare(value.AsText(), _text, _culture, CompareOptions.IgnoreCase);
                    break;
                case OperandKind.Boolean when value.Kind == CellValueKind.Boolean:
                    order = value.AsBoolean().CompareTo(_boolean);
                    break;
                default:
                    return false;
            }

            return _op switch
            {
                Operator.Less => order < 0,
                Operator.LessEqual => order <= 0,
                Operator.Greater => order > 0,
                _ => order >= 0,
            };
        }

        private bool Equal(CellValue value)
        {
            switch (_kind)
            {
                case OperandKind.Number:
                    if (value.Kind == CellValueKind.Number)
                        return value.AsNumber() == _number;
                    return _op == Operator.Equal && value.Kind == CellValueKind.Text
                        && NumberText.TryParse(value.AsText(), _culture, out var parsed) && parsed == _number;
                case OperandKind.Boolean:
                    return value.Kind == CellValueKind.Boolean && value.AsBoolean() == _boolean;
                case OperandKind.Error:
                    return value.Kind == CellValueKind.Error && value.AsError() == _error;
                default:
                    if (_text.Length == 0)
                        return _op == Operator.NotEqual
                            ? value.Kind == CellValueKind.Empty
                            : value.Kind == CellValueKind.Empty || (value.Kind == CellValueKind.Text && value.AsText().Length == 0);
                    if (value.Kind != CellValueKind.Text)
                        return false;
                    return _hasWildcards
                        ? ReferenceWildcard(_text, value.AsText(), _culture)
                        : string.Compare(value.AsText(), _text, _culture, CompareOptions.IgnoreCase) == 0;
            }
        }

        private static bool ReferenceWildcard(string pattern, string text, CultureInfo culture)
        {
            var p = culture.TextInfo.ToUpper(pattern);
            var t = culture.TextInfo.ToUpper(text);
            int pi = 0, ti = 0, starPattern = -1, starText = 0;
            while (ti < t.Length)
            {
                if (pi < p.Length && p[pi] == '*')
                {
                    starPattern = ++pi;
                    starText = ti;
                    continue;
                }

                if (pi < p.Length && Literal(p, ref pi, out var expected, out var any) && (any || expected == t[ti]))
                {
                    ti++;
                    continue;
                }

                if (starPattern < 0)
                    return false;
                pi = starPattern;
                ti = ++starText;
            }

            while (pi < p.Length && p[pi] == '*')
                pi++;
            return pi == p.Length;
        }

        private static bool Literal(string p, ref int pi, out char expected, out bool any)
        {
            expected = '\0';
            any = false;
            if (p[pi] == '?')
            {
                any = true;
                pi++;
                return true;
            }

            if (p[pi] == '~' && pi + 1 < p.Length && p[pi + 1] is '*' or '?' or '~')
            {
                expected = p[pi + 1];
                pi += 2;
                return true;
            }

            expected = p[pi];
            pi++;
            return true;
        }
    }
}
