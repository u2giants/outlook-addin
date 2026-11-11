using System;
using System.Runtime.InteropServices;

// Office COM interfaces declared locally so no Office PIA beyond Outlook is required.
namespace PopOutlookAddin
{
    public enum ext_ConnectMode { ext_cm_AfterStartup = 0, ext_cm_Startup = 1, ext_cm_External = 2, ext_cm_CommandLine = 3 }
    public enum ext_DisconnectMode { ext_dm_HostShutdown = 0, ext_dm_UserClosed = 1 }

    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object Application, ext_ConnectMode ConnectMode, [MarshalAs(UnmanagedType.IDispatch)] object AddInInst, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(ext_DisconnectMode RemoveMode, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] string GetCustomUI(string RibbonID);
    }

    [ComImport, Guid("000C0395-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonControl
    {
        [DispId(1)] string Id { get; }
        [DispId(2)] object Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
        [DispId(3)] string Tag { get; }
    }
}
