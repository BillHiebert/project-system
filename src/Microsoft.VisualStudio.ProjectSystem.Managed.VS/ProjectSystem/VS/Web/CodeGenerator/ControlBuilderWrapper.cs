// Licensed to the .NET Foundation under one or more agreements. The .NET Foundation licenses this file to you under the MIT license. See the LICENSE.md file in the project root for more information.

using System.Collections;
using System.Reflection;
using System.Web.UI;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    // Wrapper to call ControlBuilder members via private reflection
    internal static class ControlBuilderWrapper
    {
        internal static ArrayList GetSubBuilders(ControlBuilder builder)
        {
            return (ArrayList)GetProperty(builder, "SubBuilders");
        }

        internal static object GetProperty(object o, string propName)
        {
            PropertyInfo propInfo = typeof(ControlBuilder).GetProperty(propName, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            MethodInfo methodInfo = propInfo.GetGetMethod(nonPublic: true);
            return methodInfo.Invoke(o, null);
        }

        internal static ControlBuilder? GetFirstSubBuilder(ControlBuilder? builder)
        {
            ControlBuilder? firstSubBuilder = null;
            if (builder != null)
            {
                ArrayList subBuilders = GetSubBuilders(builder);
                if (subBuilders != null && subBuilders.Count > 0 && subBuilders[0] is ControlBuilder builder1)
                {
                    firstSubBuilder = builder1;
                }
            }

            return firstSubBuilder;
        }
    }
}
