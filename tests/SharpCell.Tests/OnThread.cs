using System.Runtime.ExceptionServices;

namespace SharpCell.Tests;

/// <summary>Runs test code on a thread of its own, for a stack of a known size.</summary>
internal static class OnThread
{
    /// <summary>
    /// Runs <paramref name="body"/> on a new thread and waits for it. An exception thrown inside is
    /// rethrown here, so the test fails. Left to escape the thread it would be unhandled, and the
    /// xUnit v3 runner's handler for unhandled exceptions then waits for this test, which waits in
    /// <see cref="Thread.Join()"/> for the thread: the whole run hangs.
    /// </summary>
    public static void Run(Action body, int maxStackSize)
    {
        ExceptionDispatchInfo? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                body();
            }
            catch (Exception ex)
            {
                error = ExceptionDispatchInfo.Capture(ex);
            }
        }, maxStackSize);
        thread.Start();
        thread.Join();
        error?.Throw();
    }
}
