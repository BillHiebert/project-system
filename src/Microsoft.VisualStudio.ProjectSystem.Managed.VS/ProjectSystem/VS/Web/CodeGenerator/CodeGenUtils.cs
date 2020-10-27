using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualStudio.Shell.Interop;
using IServiceProvider = System.IServiceProvider;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    // Simple struct to store information about a directive.
    internal static class CodeGenUtils
    {
        public const string WAPProjectFactoryGuid = "{349c5851-65df-11da-9384-00065b846f21}";
        public const int CLSCTX_INPROC_SERVER = 0x1;

        private static readonly Guid IID_IUnknown = new Guid("{00000000-0000-0000-C000-000000000046}");

        public static string? GetLocalRegRoot()
        {
            if (ServiceProvider.GlobalProvider.GetService(typeof(SVsShell)) is IVsShell vsShell)
            {
                vsShell.GetProperty((int)__VSSPROPID.VSSPROPID_VirtualRegistryRoot, out object obj);
                if (obj is string objValue)
                {
                    return objValue;
                }
            }

            return null;
        }
        
        public static InterfaceType? GetService<InterfaceType>(IServiceProvider serviceProvider) where InterfaceType : class
        {
            InterfaceType? service = null;

            try
            {
                if (serviceProvider != null)
                {
#pragma warning disable RS0030 // Do not used banned APIs
                    service = serviceProvider.GetService(typeof(InterfaceType)) as InterfaceType;
#pragma warning restore RS0030 // Do not used banned APIs
                }
            }
            catch
            {
            }

            return service;
        }

        public static string EnsureTrailingChar(string s, char ch)
        {
            return s.Length == 0 || s[s.Length - 1] != ch ? s + ch : s;
        }

        /// <summary>
        /// Helper to read string value from local reg using relative key name. If hkcu is true it uses
        /// the current user hive, else the local machine.
        /// </summary>
        [return: MaybeNull]
        public static ValueType GetLocalRegValue<ValueType>(bool hkcu, string key, string valueName, [AllowNull] ValueType defaultValue)
        {
            ValueType returnValue = defaultValue;
            if (!string.IsNullOrEmpty(key))
            {
                string? localreg = GetLocalRegRoot();
                if (!Strings.IsNullOrEmpty(localreg))
                {
                    string regkey = hkcu ? "HKEY_CURRENT_USER\\" : "HKEY_LOCAL_MACHINE\\";
                    regkey += EnsureTrailingChar(localreg, '\\') + key;

                    try
                    {
                        object objValue = Registry.GetValue(regkey, valueName, defaultValue);
                        if (objValue != null)
                        {
                            returnValue = (ValueType)objValue;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Returns the WAP language template guid for the base project from the registry
        /// </summary>
        public static Guid GetWAPLanguageTemplateGuid(IVsHierarchy hierarchy)
        {
            Guid languageTemplateGuid = Guid.Empty;

            if (hierarchy != null)
            {
                hierarchy.GetSite(out IOleServiceProvider? serviceProvider);
                if (serviceProvider != null)
                {
                    Guid baseProjectFactoryGuid = GetBaseProjectFactoryGuid(hierarchy);
                    if (baseProjectFactoryGuid != Guid.Empty)
                    {
                        string languageTemplates = "Projects\\" + WAPProjectFactoryGuid + "\\LanguageTemplates";
                        string? languageTemplateGuidString = GetLocalRegValue<string>(false, languageTemplates, baseProjectFactoryGuid.ToString("B"), null);
                        if (!Strings.IsNullOrEmpty(languageTemplateGuidString))
                        {
                            languageTemplateGuid = new Guid(languageTemplateGuidString);
                        }
                    }
                }
            }

            return languageTemplateGuid;
        }

        /// <summary>
        /// Returns WAP language specific property from the registry
        /// </summary>
        [return: MaybeNull]
        public static ValueType GetWAPLanguageProperty<ValueType>(IVsHierarchy hierarchy, string propertyName, [AllowNull] ValueType defaultValue)
        {
            ValueType propertyValue = defaultValue;

            if (hierarchy != null)
            {
                Guid languageTemplateGuid = GetWAPLanguageTemplateGuid(hierarchy);
                if (languageTemplateGuid != Guid.Empty)
                {
                    string languageProperties = "Projects\\" + languageTemplateGuid.ToString("B") + "\\WebApplicationProperties";
                    propertyValue = GetLocalRegValue(false, languageProperties, propertyName, defaultValue);
                }
            }

            return propertyValue;
        }

        /// <summary>
        /// Returns the language package guid from the registry
        /// </summary>
        public static Guid GetWAPLanguagePackageGuid(IVsHierarchy hierarchy)
        {
            Guid languagePackageGuid = Guid.Empty;

            if (hierarchy != null)
            {
                Guid languageTemplateGuid = GetWAPLanguageTemplateGuid(hierarchy);
                if (languageTemplateGuid != Guid.Empty)
                {
                    string languageProject = "Projects\\" + languageTemplateGuid.ToString("B");
                    string? package = GetLocalRegValue<string>(false, languageProject, "Package", null);
                    if (!Strings.IsNullOrEmpty(package))
                    {
                        languagePackageGuid = new Guid(package);
                    }
                }
            }

            return languagePackageGuid;
        }

        /// <summary>
        /// Fetches the bottom most project guid from the project properties.
        /// This should be the C#/VB/J#/Other language project factory guid. 
        /// </summary>
        public static Guid GetBaseProjectFactoryGuid(IVsHierarchy hierarchy)
        {
            Guid baseProjectFactoryGuid = Guid.Empty;

            // First try to get guid from project guid list
            if (hierarchy is IVsAggregatableProject aggregate)
            {
                try
                {
                    aggregate.GetAggregateProjectTypeGuids(out string? projectTypeGuids);
                    if (!Strings.IsNullOrEmpty(projectTypeGuids))
                    {
                        string lastGuid = projectTypeGuids.Substring(projectTypeGuids.LastIndexOf(";", StringComparison.Ordinal) + 1);
                        baseProjectFactoryGuid = new Guid(lastGuid);
                    }
                }
                catch
                {
                }
            }

            // Second try to get guid directly from hierarchy
            if (baseProjectFactoryGuid == Guid.Empty)
            {
                if (hierarchy is IPersist persist)
                {
                    try
                    {
                        persist.GetClassID(out baseProjectFactoryGuid);
                    }
                    catch
                    {
                    }
                }
            }

            return baseProjectFactoryGuid;
        }

        /// <summary>
        /// Helper to crete an instance from the local registry given a CLSID string
        /// </summary>
        public static InterfaceType? CreateInstance<InterfaceType>(string clsid) where InterfaceType : class
        {
            InterfaceType? instance = null;

            if (!string.IsNullOrEmpty(clsid))
            {
                Guid clsidGuid = Guid.Empty;
                try
                {
                    clsidGuid = new Guid(clsid);
                }
                catch
                { }

                if (clsidGuid != Guid.Empty)
                {
                    instance = CreateInstance<InterfaceType>(clsidGuid);
                }
            }

            return instance;
        }

        /// <summary>
        /// Helper to create an instance from the local registry given a CLSID Guid
        /// </summary>
        public static InterfaceType? CreateInstance<InterfaceType>(Guid clsid) where InterfaceType : class
        {
            InterfaceType? instance = null;

            if (clsid != Guid.Empty)
            {
                if (ServiceProvider.GlobalProvider.GetService(typeof(ILocalRegistry)) is ILocalRegistry localRegistry)
                {
                    IntPtr pInstance = IntPtr.Zero;
                    Guid iidUnknown = IID_IUnknown;

                    try
                    {
                        localRegistry.CreateInstance(clsid, null, ref iidUnknown, CLSCTX_INPROC_SERVER, out pInstance);
                    }
                    catch
                    { }

                    if (pInstance != IntPtr.Zero)
                    {
                        try
                        {
                            instance = Marshal.GetObjectForIUnknown(pInstance) as InterfaceType;
                        }
                        catch
                        { }

                        try
                        {
                            Marshal.Release(pInstance);
                        }
                        catch
                        { }
                    }
                }
            }

            return instance;
        }

        /// <summary>
        /// Helper to create an instance from the local registry given a CLSID Guid
        /// </summary>
        public static InterfaceType? CreateSitedInstance<InterfaceType>(IOleServiceProvider oleServiceProvider, Guid clsid) where InterfaceType : class
        {
            InterfaceType? instance = CreateInstance<InterfaceType>(clsid);

            if (instance != null)
            {
                if (instance is IObjectWithSite sitedObject)
                {
                    sitedObject.SetSite(oleServiceProvider);
                }
                else
                {
                    instance = null; // failed to site
                }
            }

            return instance;
        }

        public static string MakeRelativePath(string fullPath, string basePath)
        {
            string separator = Path.DirectorySeparatorChar.ToString();
            string tempBasePath = basePath;
            string tempFullPath = fullPath;
            string? relativePath = null;

            if (!tempBasePath.EndsWith(separator, StringComparison.Ordinal))
            {
                tempBasePath += separator;
            }

            tempFullPath = tempFullPath.ToLowerInvariant();
            tempBasePath = tempBasePath.ToLowerInvariant();

            while (!string.IsNullOrEmpty(tempBasePath))
            {
                if (tempFullPath.StartsWith(tempBasePath, StringComparison.Ordinal))
                {
                    relativePath += fullPath.Remove(0, tempBasePath.Length);
                    if (relativePath == separator)
                    {
                        relativePath = "";
                    }

                    return relativePath;
                }
                else
                {
                    tempBasePath = tempBasePath.Remove(tempBasePath.Length - 1);
                    int nLastIndex = tempBasePath.LastIndexOf(separator, StringComparison.Ordinal);
                    if (-1 != nLastIndex)
                    {
                        tempBasePath = tempBasePath.Remove(nLastIndex + 1);
                        relativePath += "..";
                        relativePath += separator;
                    }
                    else
                    {
                        return fullPath;
                    }
                }
            }

            return fullPath;
        }

        public static string? MakeRelativeUrl(string fullPath, string basePath)
        {
            string? relativeUrl = null;
            string relativePath = MakeRelativePath(fullPath, basePath);
            if (!string.IsNullOrEmpty(relativePath))
            {
                relativeUrl = relativePath.Replace(Path.DirectorySeparatorChar, '/');
            }

            return relativeUrl;
        }
    }
}
