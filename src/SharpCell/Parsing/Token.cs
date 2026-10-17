namespace SharpCell.Parsing;

internal enum TokenKind
{
    Number,
    String,
    Boolean,
    Error,
    Reference,
    Name,
    Function,
    StructuredReference,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Ampersand,
    Percent,
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
    Colon,
    Comma,
    Semicolon,
    OpenParen,
    CloseParen,
    OpenBrace,
    CloseBrace,
    At,
    Hash,
    End,
}

/// <summary>The sheet part of a reference: <c>Sheet1!</c>, or <c>Sheet1:Sheet3!</c> for a 3D reference.</summary>
internal sealed record SheetPrefix(string First, string? Last = null)
{
    /// <summary>
    /// A removed sheet, which Excel writes as <c>#REF!A1</c>. Its name is empty, which no sheet has,
    /// so the reference finds no sheet and gives <c>#REF!</c>.
    /// </summary>
    public static SheetPrefix Deleted { get; } = new("");

    /// <summary>Whether the sheet was removed.</summary>
    public bool IsDeleted => First.Length == 0;
}

internal sealed record Token(TokenKind Kind, int Start, int Length)
{
    /// <summary>Whether whitespace precedes the token; a space between operands is the intersection operator.</summary>
    public bool SpaceBefore { get; init; }

    public double Number { get; init; }

    /// <summary>String value, name, function name (as written) or raw structured reference.</summary>
    public string? Text { get; init; }

    public bool Boolean { get; init; }

    public ErrorKind Error { get; init; }

    public AreaRef Area { get; init; }

    public SheetPrefix? Sheet { get; init; }
}
