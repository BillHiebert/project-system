﻿// Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

using System.Collections.Generic;

using NuGet.SolutionRestoreManager;

namespace Microsoft.VisualStudio.ProjectSystem.VS.NuGet
{
    internal class ReferenceProperties : ImmutablePropertyCollection<IVsReferenceProperty>, IVsReferenceProperties
    {
        public ReferenceProperties(IEnumerable<IVsReferenceProperty> items)
            : base(items)
        {
        }

        protected override string GetKeyForItem(IVsReferenceProperty value) => value.Name;
    }
}
