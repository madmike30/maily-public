namespace Maily.OutlookAddin.Setup;
internal sealed class SetupScriptProgress
{
    private int applicationInstalled;
    private string activationState = "unknown";
    public bool ApplicationInstalled => Volatile.Read(ref applicationInstalled) == 1;
    public string ActivationState => Volatile.Read(ref activationState);
    public void Observe(string line)
    {
        if (line.StartsWith("MAILY_STEP|upgrade-complete|passed|", StringComparison.Ordinal)) Interlocked.Exchange(ref applicationInstalled,1);
        const string prefix = "MAILY_ACTIVATION|";
        if (!line.StartsWith(prefix,StringComparison.Ordinal)) return;
        var state=line[prefix.Length..];
        if (state is "verified" or "repair-required" or "activation-required" or "needs-support") Volatile.Write(ref activationState,state);
    }
}
