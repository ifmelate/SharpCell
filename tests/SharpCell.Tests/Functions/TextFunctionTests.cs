using System.Globalization;
using SharpCell;

namespace SharpCell.Tests.Functions;

public class TextFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public TextFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = "Hello";
        _s["A2"].Value = 12;
        _s["A3"].Value = true;
        _s["A4"].Value = CellValue.Error(ErrorKind.Div0);
        _s["B1"].Value = "a";
        _s["B3"].Value = "c";
        _s["C1"].Value = ",";
        _s["C2"].Value = ";";
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Row(params CellValue[] values)
    {
        var array = new CellValue[1, values.Length];
        for (var i = 0; i < values.Length; i++)
            array[0, i] = values[i];
        return CellValue.Array(array);
    }

    private static CellValue Column(params CellValue[] values)
    {
        var array = new CellValue[values.Length, 1];
        for (var i = 0; i < values.Length; i++)
            array[i, 0] = values[i];
        return CellValue.Array(array);
    }

    private void Check(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> Basics => new()
    {
        { "=LEN(\"Hello\")", 5 },
        { "=LEN(A2)", 2 },
        { "=LEN(A3)", 4 },
        { "=LEN(Z9)", 0 },
        { "=LEN(\"😀\")", 2 },
        { "=LEN(A4)", Err(ErrorKind.Div0) },
        { "=LENB(\"abc\")", 3 },
        { "=LEN({\"a\",\"bcd\"})", Row(1, 3) },
        { "=LEFT(\"Hello\")", "H" },
        { "=LEFT(\"Hello\",3)", "Hel" },
        { "=LEFT(\"Hello\",100)", "Hello" },
        { "=LEFT(\"Hello\",0)", "" },
        { "=LEFT(\"Hello\",)", "" },
        { "=LEFT(\"Hello\",2.9)", "He" },
        { "=LEFT(\"Hello\",-1)", Err(ErrorKind.Value) },
        { "=LEFT(A4,-1)", Err(ErrorKind.Div0) },
        { "=LEFT(12345,2)", "12" },
        { "=RIGHT(\"Hello\")", "o" },
        { "=RIGHT(\"Hello\",3)", "llo" },
        { "=RIGHT(\"Hello\",9)", "Hello" },
        { "=RIGHT(\"Hello\",-2)", Err(ErrorKind.Value) },
        { "=RIGHTB(\"Hello\",2)", "lo" },
        { "=MID(\"Hello\",2,3)", "ell" },
        { "=MID(\"Hello\",4,10)", "lo" },
        { "=MID(\"Hello\",6,1)", "" },
        { "=MID(\"Hello\",0,1)", Err(ErrorKind.Value) },
        { "=MID(\"Hello\",1,-1)", Err(ErrorKind.Value) },
        { "=MID(\"Hello\",1/0,1)", Err(ErrorKind.Div0) },
        { "=MID(\"Hello\",{1,2,3},2)", Row("He", "el", "ll") },
        { "=REPLACE(\"abcdef\",2,3,\"X\")", "aXef" },
        { "=REPLACE(\"abc\",10,1,\"X\")", "abcX" },
        { "=REPLACE(\"abc\",1,0,\"X\")", "Xabc" },
        { "=REPLACE(\"abc\",0,1,\"X\")", Err(ErrorKind.Value) },
        { "=REPLACE(\"abc\",\"NaN\",1,\"x\")", Err(ErrorKind.Value) },
        { "=REPT(\"ab\",3)", "ababab" },
        { "=REPT(\"ab\",0)", "" },
        { "=REPT(\"ab\",2.9)", "abab" },
        { "=REPT(\"ab\",-1)", Err(ErrorKind.Value) },
        { "=REPT(\"ab\",20000)", Err(ErrorKind.Value) },
        { "=REPT(\"\",100000)", "" },
        { "=REPT(TRUE,2)", "TRUETRUE" },
    };

    [Theory]
    [MemberData(nameof(Basics))]
    public void Length_and_slicing(string formula, CellValue expected) => Check(formula, expected);

    public static TheoryData<string, CellValue> Searching => new()
    {
        { "=FIND(\"l\",\"Hello\")", 3 },
        { "=FIND(\"l\",\"Hello\",4)", 4 },
        { "=FIND(\"L\",\"Hello\")", Err(ErrorKind.Value) },
        { "=FIND(\"\",\"Hello\")", 1 },
        { "=FIND(\"\",\"Hello\",6)", 6 },
        { "=FIND(\"\",\"Hello\",7)", Err(ErrorKind.Value) },
        { "=FIND(\"l\",\"Hello\",0)", Err(ErrorKind.Value) },
        { "=FIND(\"*\",\"a*b\")", 2 },
        { "=FINDB(\"b\",\"abc\")", 2 },
        { "=SEARCH(\"L\",\"Hello\")", 3 },
        { "=SEARCH(\"l?o\",\"Hello\")", 3 },
        { "=SEARCH(\"h*o\",\"ahello\")", 2 },
        { "=SEARCH(\"*c\",\"abc\")", 1 },
        { "=SEARCH(\"~*\",\"ab*c\")", 3 },
        { "=SEARCH(\"~~\",\"Hi ~Bang!\")", 4 },
        { "=SEARCH(\"~\",\"Hello\")", 1 },
        { "=SEARCH(\"~d\",\"My dear\")", 4 },
        { "=SEARCH(\"~~*\",\"Ho*la\")", Err(ErrorKind.Value) },
        { "=SEARCH(\"x\",\"Hello\")", Err(ErrorKind.Value) },
        { "=SEARCH(\"l\",\"Hello\",5)", Err(ErrorKind.Value) },
        { "=SEARCH(\"\",\"Hello\",3)", 3 },
        { "=SEARCH({\"e\",\"o\"},\"Hello\")", Row(2, 5) },
        { "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "a+b+c" },
        { "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "a-b+c" },
        { "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",3)", "a-b-c" },
        { "=SUBSTITUTE(\"aaa\",\"aa\",\"b\")", "ba" },
        { "=SUBSTITUTE(\"aaaa\",\"aa\",\"b\",2)", "aab" },
        { "=SUBSTITUTE(\"abc\",\"\",\"x\")", "abc" },
        { "=SUBSTITUTE(\"abc\",\"B\",\"x\")", "abc" },
        { "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",0)", Err(ErrorKind.Value) },
        { "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",\"NaN\")", Err(ErrorKind.Value) },
        { "=EXACT(\"a\",\"a\")", true },
        { "=EXACT(\"a\",\"A\")", false },
        { "=EXACT(12,\"12\")", true },
        { "=EXACT(A4,\"x\")", Err(ErrorKind.Div0) },
    };

    [Theory]
    [MemberData(nameof(Searching))]
    public void Finding_and_replacing(string formula, CellValue expected) => Check(formula, expected);

    public static TheoryData<string, CellValue> Casing => new()
    {
        { "=UPPER(\"Straße 1\")", "STRAßE 1" },
        { "=LOWER(\"ÀB\")", "àb" },
        { "=PROPER(\"hello wORLD\")", "Hello World" },
        { "=PROPER(\"2nd don't\")", "2Nd Don'T" },
        { "=PROPER(\"三体\")", "三体" },
        { "=TRIM(\"  a   b  \")", "a b" },
        { "=TRIM(\"a  b\")", "a  b" },
        { "=CLEAN(\"a\u0001b\tc\u007F\")", "abc\u007F" },
        { "=UPPER(A4)", Err(ErrorKind.Div0) },
        { "=CHAR(65)", "A" },
        { "=CHAR(128)", "€" },
        { "=CHAR(129)", "\u0081" },
        { "=CHAR(255)", "ÿ" },
        { "=CHAR(65.9)", "A" },
        { "=CHAR(0)", Err(ErrorKind.Value) },
        { "=CHAR(256)", Err(ErrorKind.Value) },
        { "=CODE(\"A\")", 65 },
        { "=CODE(\"€uro\")", 128 },
        { "=CODE(\"é\")", 233 },
        { "=CODE(\"の\")", 63 },
        { "=CODE(\"\")", Err(ErrorKind.Value) },
        { "=UNICHAR(9731)", "☃" },
        { "=UNICHAR(128512)", "😀" },
        { "=UNICHAR(0)", Err(ErrorKind.Value) },
        { "=UNICHAR(1114112)", Err(ErrorKind.Value) },
        { "=UNICHAR(55296)", Err(ErrorKind.NA) },
        { "=UNICODE(\"の\")", 12398 },
        { "=UNICODE(\"😀\")", 128512 },
        { "=UNICODE(1)", 49 },
        { "=UNICODE(\"\")", Err(ErrorKind.Value) },
        { "=UNICODE(Z9)", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(Casing))]
    public void Case_and_characters(string formula, CellValue expected) => Check(formula, expected);

    public static TheoryData<string, CellValue> Conversions => new()
    {
        { "=T(\"a\")", "a" },
        { "=T(12)", "" },
        { "=T(A3)", "" },
        { "=T(A4)", Err(ErrorKind.Div0) },
        { "=VALUE(\"12\")", 12 },
        { "=VALUE(\" 1,234.5 \")", 1234.5 },
        { "=VALUE(\"40%\")", 0.4 },
        { "=VALUE(\"1e-5\")", 1e-5 },
        { "=VALUE(12)", 12 },
        { "=VALUE(Z9)", 0 },
        { "=VALUE(\"\")", Err(ErrorKind.Value) },
        { "=VALUE(\"abc\")", Err(ErrorKind.Value) },
        { "=VALUE(TRUE)", Err(ErrorKind.Value) },
        { "=VALUE(A4)", Err(ErrorKind.Div0) },
        { "=NUMBERVALUE(\"2.500,27\",\",\",\".\")", 2500.27 },
        { "=NUMBERVALUE(\"3.5%\")", 0.035 },
        { "=NUMBERVALUE(\"9%%\")", 0.0009 },
        { "=NUMBERVALUE(\" 1 2 3 \")", 123 },
        { "=NUMBERVALUE(\"\")", 0 },
        { "=NUMBERVALUE(\"1,2,3\")", 123 },
        { "=NUMBERVALUE(\"1.2.3\")", Err(ErrorKind.Value) },
        { "=NUMBERVALUE(\"1.5,0\")", Err(ErrorKind.Value) },
        { "=NUMBERVALUE(\"1,5\",\",\",\",\")", Err(ErrorKind.Value) },
        { "=NUMBERVALUE(\"-1.5E2\")", -150 },
        { "=NUMBERVALUE(\"abc\")", Err(ErrorKind.Value) },
        { "=VALUETOTEXT(\"a\")", "a" },
        { "=VALUETOTEXT(\"a\",1)", "\"a\"" },
        { "=VALUETOTEXT(12.5,1)", "12.5" },
        { "=VALUETOTEXT(TRUE)", "TRUE" },
        { "=VALUETOTEXT(A4)", "#DIV/0!" },
        { "=VALUETOTEXT(Z9)", "" },
        { "=VALUETOTEXT(1,2)", Err(ErrorKind.Value) },
        { "=ARRAYTOTEXT({1,\"a\";TRUE,#N/A})", "1, a, TRUE, #N/A" },
        { "=ARRAYTOTEXT({1,\"a\";TRUE,#N/A},1)", "{1,\"a\";TRUE,#N/A}" },
        { "=ARRAYTOTEXT(A1:B1)", "Hello, a" },
        { "=ARRAYTOTEXT(A1:A4,1)", "{\"Hello\";12;TRUE;#DIV/0!}" },
        { "=ARRAYTOTEXT(5)", "5" },
        { "=ARRAYTOTEXT(A1:B1,3)", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(Conversions))]
    public void Text_and_number_conversions(string formula, CellValue expected) => Check(formula, expected);

    public static TheoryData<string, CellValue> Joining => new()
    {
        { "=CONCATENATE(\"a\",1,TRUE)", "a1TRUE" },
        { "=CONCATENATE(\"a\",,\"b\")", "ab" },
        { "=CONCATENATE(\"a\",A4)", Err(ErrorKind.Div0) },
        { "=CONCATENATE({\"a\",\"b\"},\"!\")", Row("a!", "b!") },
        { "=CONCAT(A1:B3)", "Helloa12TRUEc" },
        { "=CONCAT({1,2;3,4},\"x\")", "1234x" },
        { "=CONCAT(A1:A4)", Err(ErrorKind.Div0) },
        { "=CONCAT(Z1:Z9)", "" },
        { "=CONCAT(REPT(\"a\",20000),REPT(\"b\",20000))", Err(ErrorKind.Value) },
        { "=TEXTJOIN(\",\",TRUE,A1:B3)", "Hello,a,12,TRUE,c" },
        { "=TEXTJOIN(\",\",FALSE,A1:B3)", "Hello,a,12,,TRUE,c" },
        { "=TEXTJOIN(\"-\",TRUE,\"a\",\"\",\"b\")", "a-b" },
        { "=TEXTJOIN(\"-\",FALSE,\"a\",\"\",\"b\")", "a--b" },
        { "=TEXTJOIN(C1:C2,TRUE,\"a\",\"b\",\"c\",\"d\")", "a,b;c,d" },
        { "=TEXTJOIN(\"\",FALSE,A1:B1)", "Helloa" },
        { "=TEXTJOIN(\",\",TRUE,A4)", Err(ErrorKind.Div0) },
        { "=TEXTJOIN(\",\",\"x\",\"a\")", Err(ErrorKind.Value) },
        { "=TEXTJOIN(\",\",TRUE,Z:Z)", "" },
        { "=TEXTJOIN(\",\",FALSE,Z:Z)", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(Joining))]
    public void Joining_text(string formula, CellValue expected) => Check(formula, expected);

    public static TheoryData<string, CellValue> Formatting => new()
    {
        { "=TEXT(1234.567,\"#,##0.00\")", "1,234.57" },
        { "=TEXT(\"12\",\"0.00\")", "12.00" },
        { "=TEXT(\"abc\",\"0.00\")", "abc" },
        { "=TEXT(\"abc\",\"\"\"[\"\"@\"\"]\"\"\")", "[abc]" },
        { "=TEXT(TRUE,\"0\")", "TRUE" },
        { "=TEXT(Z9,\"0.00\")", "0.00" },
        { "=TEXT(0.5,\"h:mm AM/PM\")", "12:00 PM" },
        { "=TEXT(-1,\"d/m/yyyy\")", Err(ErrorKind.Value) },
        { "=TEXT(1,\"mm###\")", Err(ErrorKind.Value) },
        { "=TEXT(A4,\"0\")", Err(ErrorKind.Div0) },
        { "=TEXT(1,A4)", Err(ErrorKind.Div0) },
        { "=TEXT(12,0)", "12" },
        { "=TEXT({1,2},\"00\")", Row("01", "02") },
        { "=FIXED(1234.567)", "1,234.57" },
        { "=FIXED(1234.567,1)", "1,234.6" },
        { "=FIXED(1234.567,-1)", "1,230" },
        { "=FIXED(-1234.567,-1,TRUE)", "-1230" },
        { "=FIXED(1234.567,0,TRUE)", "1235" },
        { "=FIXED(2.675,2)", "2.68" },
        { "=FIXED(1,128)", Err(ErrorKind.Value) },
        { "=FIXED(\"x\")", Err(ErrorKind.Value) },
        { "=DOLLAR(1234.567)", "$1,234.57" },
        { "=DOLLAR(-1234.567,-2)", "($1,200)" },
        { "=DOLLAR(0.5,0)", "$1" },
        { "=DOLLAR(1,128)", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(Formatting))]
    public void Number_formatting(string formula, CellValue expected) => Check(formula, expected);

    [Fact]
    public void Numbers_and_dollars_follow_the_workbook_culture()
    {
        _wb.Culture = CultureInfo.GetCultureInfo("de-DE");
        Check("=FIXED(1234.567)", "1.234,57");
        Check("=TEXT(1234.5,\"#,##0.00\")", "1.234,50");
        Check("=NUMBERVALUE(\"1.234,5\")", 1234.5);
        Check("=DOLLAR(1234.567)", "1.234,57 €");
        Check("=DOLLAR(-1234.567)", "-1.234,57 €");
    }

    [Fact]
    public void Ranges_lift_scalar_parameters() =>
        Check("=LEFT(A1:A2,{1,2})", CellValue.Array(new CellValue[,] { { "H", "He" }, { "1", "12" } }));

    [Fact]
    public void Text_over_the_limit_is_an_error()
    {
        Check("=LEN(REPT(\"a\",32767))", 32767);
        Check("=REPT(\"a\",32768)", Err(ErrorKind.Value));
        Check("=REPLACE(REPT(\"a\",32767),1,0,\"b\")", Err(ErrorKind.Value));
        Check("=SUBSTITUTE(REPT(\"a\",20000),\"a\",\"bb\")", Err(ErrorKind.Value));
        Check("=CONCATENATE(REPT(\"a\",32767),\"b\")", Err(ErrorKind.Value));
    }

    [Fact]
    public void Column_results_spill() => Check("=UPPER(A1:A2)", Column("HELLO", "12"));
}
