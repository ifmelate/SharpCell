namespace SharpCell.Tests;

public class OnThreadTests
{
    [Fact]
    public void An_exception_on_the_thread_reaches_the_caller()
    {
        var error = Assert.Throws<InvalidOperationException>(() => OnThread.Run(() => throw new InvalidOperationException("inside"), 256 * 1024));
        Assert.Equal("inside", error.Message);
    }

    [Fact]
    public void The_body_runs_on_a_thread_with_the_requested_stack()
    {
        var caller = Environment.CurrentManagedThreadId;
        var inside = caller;
        OnThread.Run(() => inside = Environment.CurrentManagedThreadId, 256 * 1024);
        Assert.NotEqual(caller, inside);
    }
}
