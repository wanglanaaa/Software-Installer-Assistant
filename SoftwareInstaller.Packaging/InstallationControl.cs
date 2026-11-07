namespace SoftwareInstaller.Packaging;

public sealed class InstallationControl : IDisposable
{
    private readonly CancellationTokenSource stopWaiting = new();
    private int terminateRequested;
    internal CancellationToken Token => stopWaiting.Token;
    internal bool TerminationRequested => Volatile.Read(ref terminateRequested) != 0;

    public void StopWaiting() => stopWaiting.Cancel();
    public void TerminateCurrent()
    {
        Interlocked.Exchange(ref terminateRequested, 1);
        stopWaiting.Cancel();
    }

    public void Dispose() => stopWaiting.Dispose();
}
