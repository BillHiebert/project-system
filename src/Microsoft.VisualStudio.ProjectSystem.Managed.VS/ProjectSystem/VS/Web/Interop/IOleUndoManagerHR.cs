// Licensed to the .NET Foundation under one or more agreements. The .NET Foundation licenses this file to you under the MIT license. See the LICENSE.md file in the project root for more information.

using System.Runtime.InteropServices;
using Microsoft.VisualStudio.OLE.Interop;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    [ComImport]
    [Guid("D001F200-EF97-11CE-9BC9-00AA00608E01")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleUndoManagerHR
    {
        [PreserveSig]
        int Open(
            [In][MarshalAs(UnmanagedType.Interface)] IOleParentUndoUnit pPUU);

        [PreserveSig]
        int Close(
            [In][MarshalAs(UnmanagedType.Interface)] IOleParentUndoUnit pPUU,
            [In][ComAliasName("Microsoft.VisualStudio.OLE.Interop.BOOL")] int fCommit);

        [PreserveSig]
        int Add(
            [In][MarshalAs(UnmanagedType.Interface)] IOleUndoUnit pUU);

        [PreserveSig]
        int GetOpenParentState(
            [Out][ComAliasName("Microsoft.VisualStudio.OLE.Interop.DWORD")] out uint pdwState);

        [PreserveSig]
        int DiscardFrom(
            [In][MarshalAs(UnmanagedType.Interface)] IOleUndoUnit pUU);

        [PreserveSig]
        int UndoTo(
            [In][MarshalAs(UnmanagedType.Interface)] IOleUndoUnit pUU);

        [PreserveSig]
        int RedoTo(
            [In][MarshalAs(UnmanagedType.Interface)] IOleUndoUnit pUU);

        [PreserveSig]
        int EnumUndoable(
            [Out][MarshalAs(UnmanagedType.Interface)] out IEnumOleUndoUnits ppEnum);

        [PreserveSig]
        int EnumRedoable([MarshalAs(UnmanagedType.Interface)] out IEnumOleUndoUnits ppEnum);

        [PreserveSig]
        int GetLastUndoDescription(
            [Out][MarshalAs(UnmanagedType.BStr)] out string pBstr);

        [PreserveSig]
        int GetLastRedoDescription(
            [Out][MarshalAs(UnmanagedType.BStr)] out string pBstr);

        [PreserveSig]
        int Enable(
            [In][ComAliasName("Microsoft.VisualStudio.OLE.Interop.BOOL")] int fEnable);
    }
}
