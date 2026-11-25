using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// COM's message filter, so that a call Excel rejects while it is busy (refreshing pivot
/// tables when a file opens, for example) is tried again instead of failing with
/// RPC_E_CALL_REJECTED. Registered on the STA thread that talks to Excel.
/// </summary>
[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);

    [PreserveSig]
    int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
}

/// <summary>Retries rejected calls for up to a minute, a quarter of a second apart.</summary>
[SupportedOSPlatform("windows")]
internal sealed class MessageFilter : IOleMessageFilter
{
    private const int Handled = 0;            // SERVERCALL_ISHANDLED
    private const int WaitDefault = 2;        // PENDINGMSG_WAITDEFPROCESS
    private const int Cancel = -1;
    private const int RetryAfterMilliseconds = 250;
    private const int GiveUpAfterMilliseconds = 60_000;

    /// <summary>Registers the filter on the current thread.</summary>
    public static void Register()
    {
        _ = CoRegisterMessageFilter(new MessageFilter(), out _);
    }

    /// <summary>Puts the thread's default behaviour back.</summary>
    public static void Revoke()
    {
        _ = CoRegisterMessageFilter(null, out _);
    }

    public int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo)
    {
        return Handled;
    }

    public int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType)
    {
        if (tickCount < GiveUpAfterMilliseconds)
        {
            return RetryAfterMilliseconds;
        }

        return Cancel;
    }

    public int MessagePending(IntPtr taskCallee, int tickCount, int pendingType)
    {
        return WaitDefault;
    }

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
}
