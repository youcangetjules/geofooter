using System;
using System.Runtime.InteropServices;

namespace Extensibility
{
    // IDTExtensibility2 is a dual interface — Outlook calls OnConnection via vtable.
    // InterfaceIsIDispatch would leave the vtable empty past IDispatch and crash on load.
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [In, MarshalAs(UnmanagedType.IDispatch)] object Application,
            [In] ext_ConnectMode ConnectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] ext_DisconnectMode RemoveMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5,
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3,
    }
}

namespace Microsoft.Office.Core
{
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI(string RibbonID);
    }

    [ComImport]
    [Guid("000C03A7-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonUI
    {
        [DispId(1)]
        void Invalidate();

        [DispId(2)]
        void InvalidateControl(string ControlID);

        [DispId(3)]
        void InvalidateControlMso(string ControlID);

        [DispId(4)]
        void ActivateTab(string ControlID);

        [DispId(5)]
        void ActivateTabMso(string ControlID);

        [DispId(6)]
        void ActivateTabQ(string ControlID, string Namespace);
    }

    [ComImport]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonControl
    {
        [DispId(1)]
        string Id
        {
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
        }

        [DispId(2)]
        object Context
        {
            [return: MarshalAs(UnmanagedType.IDispatch)]
            get;
        }

        [DispId(3)]
        string Tag
        {
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
        }
    }
}

namespace stdole
{
    [ComImport]
    [Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPictureDisp
    {
    }
}
