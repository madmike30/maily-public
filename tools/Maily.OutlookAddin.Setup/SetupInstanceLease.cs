using System.Security.Principal;

namespace Maily.OutlookAddin.Setup;

internal sealed class SetupInstanceLease : IDisposable
{
    private Mutex? mutex;
    private SetupInstanceLease(Mutex mutex) => this.mutex = mutex;

    public static string CurrentUserName()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var sid = identity.User?.Value ?? throw new InvalidOperationException("Unable to identify the setup user.");
        return @"Global\Maily.OutlookAddin.Setup." + sid;
    }

    public static SetupInstanceLease? TryAcquire(string name)
    {
        var mutex = new Mutex(false, name);
        try
        {
            bool acquired;
            try { acquired = mutex.WaitOne(0); }
            catch (AbandonedMutexException) { acquired = true; }
            if (acquired) return new SetupInstanceLease(mutex);
            mutex.Dispose();
            return null;
        }
        catch { mutex.Dispose(); throw; }
    }

    public void Dispose()
    {
        var owned = Interlocked.Exchange(ref mutex, null);
        if (owned is null) return;
        try { owned.ReleaseMutex(); }
        finally { owned.Dispose(); }
    }
}
