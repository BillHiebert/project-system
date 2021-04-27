﻿// Licensed to the .NET Foundation under one or more agreements. The .NET Foundation licenses this file to you under the MIT license. See the LICENSE.md file in the project root for more information.

using System;
using System.Threading.Tasks;
using Moq;
using IOleAsyncServiceProvider = Microsoft.VisualStudio.Shell.IAsyncServiceProvider;

namespace Microsoft.VisualStudio.Shell.Interop
{
    internal static class IOleAsyncServiceProviderFactory
    {
        public static IOleAsyncServiceProvider ImplementQueryServiceAsync(object? service, Type clsid)
        {
            var mock = new Mock<IOleAsyncServiceProvider>();

            mock.Setup(p => p.GetServiceAsync(clsid))
              .Returns((Task<object?>)IVsTaskFactory.FromResult(service));

            return mock.Object;
        }
    }
}
