// Licensed to the .NET Foundation under one or more agreements. The .NET Foundation licenses this file to you under the MIT license. See the LICENSE.md file in the project root for more information.

using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell.Interop;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    [ComImport()]
    [Guid("34a96733-b207-4303-84cc-cf4cccf5680c")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IVsCodeBehindCodeGeneratorService
    {
        [PreserveSig]
        int InitGenerator(
            [In][MarshalAs(UnmanagedType.Interface)] IVsHierarchy hierarchy,
            [In][MarshalAs(UnmanagedType.LPWStr)] string appVirtualPath,
            [In][MarshalAs(UnmanagedType.LPWStr)] string appPhysicalPath);

        [PreserveSig]
        int CloseGenerator();

        [PreserveSig]
        int UpdateDesignerClass(
            [In][MarshalAs(UnmanagedType.LPWStr)] string document,
            [In][MarshalAs(UnmanagedType.LPWStr)] string codeBehind,
            [In][MarshalAs(UnmanagedType.LPWStr)] string codeBehindFile,
            [In][MarshalAs(UnmanagedType.Bool)] bool force,
            [In][MarshalAs(UnmanagedType.Bool)] bool create);
    }
}
