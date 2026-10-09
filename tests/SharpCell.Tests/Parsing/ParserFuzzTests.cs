using System.Diagnostics;
using System.Text;
using SharpCell;
using SharpCell.Parsing;

namespace SharpCell.Tests.Parsing;

/// <summary>
/// Generated hostile input. The parser may accept or reject it, but may only throw
/// <see cref="FormulaParseException"/>, must not hang or overflow the stack, and whatever it
/// accepts must print to text that parses back to the same tree, up to parentheses.
/// </summary>
public class ParserFuzzTests
{
    // Set SHARPCELL_FUZZ_ITERATIONS for a longer local run; CI uses the default.
    private static readonly int Iterations =
        int.TryParse(Environment.GetEnvironmentVariable("SHARPCELL_FUZZ_ITERATIONS"), out var n) ? n : 3000;
    // Wall-clock time, so it catches hangs and exponential blowups, not slow machines: an input that
    // parses in a millisecond can wait over a second for a CPU when both target frameworks run at once.
    // SHARPCELL_FUZZ_BUDGET_MS overrides it.
    private static readonly TimeSpan PerInputBudget = TimeSpan.FromMilliseconds(
        int.TryParse(Environment.GetEnvironmentVariable("SHARPCELL_FUZZ_BUDGET_MS"), out var ms) ? ms : 10_000);

    private static readonly string[] TokenPool =
    [
        "1", "0.5", "1E+3", "1e", ".", "\"s\"", "\"", "\"\"", "TRUE", "FALSE", "#N/A", "#REF!", "#", "#DIV/0!",
        "A1", "$B$2", "A:A", "1:1", "$", "XFD1048576", "XFE1", "Sheet1!", "'My Sheet'!", "'", "!", "Sheet1:Sheet2!",
        "SUM(", "IF(", "LAMBDA(", "LET(", "_xlfn.", "_xlpm.x", "_xlfn.ANCHORARRAY(", "_xlfn.SINGLE(",
        "(", ")", "{", "}", ",", ";", "+", "-", "*", "/", "^", "&", "%", "=", "<>", "<=", ">=", "<", ">", "@", ":",
        " ", "  ", "\n", "\t", " ", "Name", "Налог", "Table1[", "[", "]", "[@Col]", "Sales[[#Headers],[#Data]]", "[@[Col A]]", "T['[x']]", "Sales[[A]:[B]]", "Sales[#This Row]", "[1]", "R[1]C", "RC", "R1C1",
        "€", "😀", "\0", "\uD800", "\uDFFF", "\\",
    ];

    private static readonly string[] Seeds =
    [
        "SUM(A1:B10)*2", "IF(A1>0,\"pos\",IF(A1<0,\"neg\",\"zero\"))", "IFERROR(VLOOKUP($A2,Data!$A:$D,4,FALSE),0)",
        "LET(x,A1*2,LAMBDA(y,x+y)(3))", "SUM((A1,B1),C1)", "A1:B2 B1:C3", "{1,2;3,4}*-A1#", "@A1:A10&\"x\"",
        "'My Sheet'!$A$1+Sheet1:Sheet3!B2", "-2^-2%", "Table1[[#This Row],[Col]]+1",
    ];

    public enum Generator
    {
        TokenSoup,
        RandomCharacters,
        DeepNesting,
        Mutation,
        LongChain,
        PostfixChain,
    }

    [Theory]
    [InlineData(Generator.TokenSoup, 1)]
    [InlineData(Generator.TokenSoup, 2)]
    [InlineData(Generator.RandomCharacters, 3)]
    [InlineData(Generator.DeepNesting, 4)]
    [InlineData(Generator.Mutation, 5)]
    [InlineData(Generator.Mutation, 6)]
    [InlineData(Generator.LongChain, 7)]
    [InlineData(Generator.PostfixChain, 8)]
    public void Parser_survives_generated_input(Generator generator, int seed)
    {
        var failures = new List<string>();
        OnThread.Run(() =>
        {
            var random = new Random(seed);
            var iterations = generator is Generator.DeepNesting or Generator.LongChain or Generator.PostfixChain
                ? Iterations / 10
                : Iterations;
            for (var i = 0; i < iterations && failures.Count < 10; i++)
            {
                var text = Generate(generator, random);
                var origin = new CellAddress(random.Next(1, CellAddress.MaxRow + 1), random.Next(1, CellAddress.MaxColumn + 1));
                var style = random.Next(4) == 0 ? ReferenceStyle.R1C1 : ReferenceStyle.A1;
                var failure = Check(text, origin, style);
                if (failure is not null)
                    failures.Add($"{failure}\n  input: {Escape(text)}\n  origin: {origin}, style: {style}");
            }
        }, maxStackSize: 1024 * 1024);

        Assert.True(failures.Count == 0, string.Join("\n\n", failures));
    }

    private static string? Check(string text, CellAddress origin, ReferenceStyle style)
    {
        var stopwatch = Stopwatch.StartNew();
        var failure = Verify(text, origin, style);
        return failure ?? (stopwatch.Elapsed > PerInputBudget ? $"took {stopwatch.Elapsed.TotalMilliseconds:0} ms" : null);
    }

    private static string? Verify(string text, CellAddress origin, ReferenceStyle style)
    {
        try
        {
            FormulaNode node;
            try
            {
                node = FormulaParser.Parse(text, origin, style);
            }
            catch (FormulaParseException ex)
            {
                return ex.Position < 0 || ex.Position > text.Length ? $"position {ex.Position} outside the text" : null;
            }

            var printed = FormulaPrinter.Print(node, origin, style);
            if (printed.Length > FormulaLimits.MaxLength)
                return null;

            FormulaNode reparsed;
            try
            {
                reparsed = FormulaParser.Parse(printed, origin, style);
            }
            catch (FormulaParseException ex)
            {
                return $"printed text does not parse ({ex.Message} at {ex.Position}): {Escape(printed)}";
            }

            // The printer may add parentheses (an ANCHORARRAY argument becomes "(x+1)#"); they must not change meaning.
            if (TreeDump.Dump(reparsed, ignoreParentheses: true) != TreeDump.Dump(node, ignoreParentheses: true))
                return $"printed text parses to a different tree: {Escape(printed)}";

            return null;
        }
        catch (Exception ex)
        {
            return $"{ex.GetType().Name}: {ex.Message}";
        }
    }

    private static string Generate(Generator generator, Random random) => generator switch
    {
        Generator.TokenSoup => TokenSoup(random, random.Next(1, 40)),
        Generator.RandomCharacters => RandomCharacters(random),
        Generator.DeepNesting => DeepNesting(random),
        Generator.Mutation => Mutate(random, Seeds[random.Next(Seeds.Length)]),
        Generator.LongChain => LongChain(random),
        _ => PostfixChain(random),
    };

    private static string TokenSoup(Random random, int count)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < count; i++)
            sb.Append(TokenPool[random.Next(TokenPool.Length)]);
        return sb.ToString();
    }

    private static string RandomCharacters(Random random)
    {
        var length = random.Next(0, 60);
        var sb = new StringBuilder(length);
        for (var i = 0; i < length; i++)
        {
            sb.Append(random.Next(4) switch
            {
                0 => (char)random.Next(0, 32),
                1 => (char)random.Next(32, 127),
                2 => (char)random.Next(0xD800, 0xE000),
                _ => (char)random.Next(127, 0x10000),
            });
        }

        return sb.ToString();
    }

    private static string DeepNesting(Random random)
    {
        string[] openers = ["(", "SUM(", "-", "@", "{", "LAMBDA(x,", "+", "A1:("];
        var depth = random.Next(50, FormulaLimits.MaxLength);
        var sb = new StringBuilder();
        for (var i = 0; i < depth && sb.Length < FormulaLimits.MaxLength - 10; i++)
            sb.Append(openers[random.Next(openers.Length)]);
        sb.Append('1');
        for (var i = 0; i < depth && sb.Length < FormulaLimits.MaxLength; i++)
            sb.Append(random.Next(3) == 0 ? "}" : ")");
        return sb.ToString();
    }

    private static string Mutate(Random random, string seed)
    {
        var chars = new List<char>(seed);
        var mutations = random.Next(1, 6);
        for (var m = 0; m < mutations; m++)
        {
            var position = random.Next(chars.Count + 1);
            switch (random.Next(3))
            {
                case 0:
                    chars.InsertRange(position, TokenPool[random.Next(TokenPool.Length)]);
                    break;
                case 1 when chars.Count > 0:
                    chars.RemoveAt(Math.Min(position, chars.Count - 1));
                    break;
                default:
                    if (chars.Count > 0)
                        chars[Math.Min(position, chars.Count - 1)] = TokenPool[random.Next(TokenPool.Length)][0];
                    break;
            }
        }

        return new string(chars.ToArray());
    }

    private static string LongChain(Random random)
    {
        string[] operators = ["+", "-", "*", "/", "^", "&", "=", ":", " ", ","];
        string[] operands = ["A1", "1", "\"x\"", "B2:C3", "Name", "(1)", "-1", "SUM(1)"];
        var sb = new StringBuilder(operands[random.Next(operands.Length)]);
        while (sb.Length < FormulaLimits.MaxLength - random.Next(0, 30))
            sb.Append(operators[random.Next(operators.Length)]).Append(operands[random.Next(operands.Length)]);
        return sb.ToString();
    }

    // Postfix operators grow the tree without nesting the text: 1%%%, A1###, f(1)(1)(1).
    // Tails come from one group so the chain stays parsable and reaches the printer.
    private static string PostfixChain(Random random)
    {
        string[] heads = ["1", "A1", "(1)", "SUM(1)", "LAMBDA(x,x)", "-1", "@A1"];
        string[][] groups = [["%"], ["#"], ["(1)", "()", "(A1)"]];
        var group = groups[random.Next(groups.Length)];
        var sb = new StringBuilder(heads[random.Next(heads.Length)]);
        var count = random.Next(10, FormulaLimits.MaxLength);
        for (var i = 0; i < count; i++)
        {
            var tail = group[random.Next(group.Length)];
            if (sb.Length + tail.Length > FormulaLimits.MaxLength)
                break;
            sb.Append(tail);
        }

        return sb.ToString();
    }

    private static string Escape(string text)
    {
        var sb = new StringBuilder();
        foreach (var c in text.Length > 300 ? text[..300] + "…" : text)
            sb.Append(c < 32 || char.IsSurrogate(c) ? $"\\u{(int)c:X4}" : c.ToString());
        return sb.ToString();
    }
}
