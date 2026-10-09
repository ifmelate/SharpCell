using System;

namespace SharpCell.Docs;

/// <summary>A problem in the documentation sources, reported to the author without a stack trace.</summary>
internal sealed class DocsException(string message) : Exception(message);
