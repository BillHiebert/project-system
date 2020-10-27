
using EnvDTE;
using EnvDTE80;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using VSLangProj;
using IServiceProvider = System.IServiceProvider;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;
using Microsoft.VisualStudio.Shell.Flavor;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{

    internal class VsHierarchyItem
    {
        private readonly uint _vsitemid;
        private readonly IVsHierarchy _hier;
        private IServiceProvider? _serviceProvider;
        private int _isFile = -1;

        internal VsHierarchyItem(uint id, IVsHierarchy hier)
        {
            if (hier == null)
            {
                throw new ArgumentNullException(nameof(hier));
            }

            _vsitemid = id;
            _hier = hier;
        }

        internal VsHierarchyItem(IVsHierarchy hier)
        {
            if (hier == null)
            {
                throw new ArgumentNullException(nameof(hier));
            }

            _vsitemid = VSConstants.VSITEMID_ROOT;
            _hier = hier;
        }

        /// <summary>
        /// Locates the item in the provided hierarchy using the provided moniker
        /// and return a VsHierarchyItem for it
        /// NOTE!!! This ONLY works for files!
        /// </summary>
        internal static VsHierarchyItem? CreateFromMoniker(string moniker, IVsHierarchy hier)
        {
            VsHierarchyItem? item = null;

            if (!string.IsNullOrEmpty(moniker) && hier != null)
            {
                if (hier is IVsProject proj)
                {
                    VSDOCUMENTPRIORITY[] priority = new VSDOCUMENTPRIORITY[1];
                    int hr = proj.IsDocumentInProject(moniker, out int isFound, priority, out uint itemid);
                    if (ErrorHandler.Succeeded(hr) && isFound != 0 && itemid != VSConstants.VSITEMID_NIL)
                    {
                        item = new VsHierarchyItem(itemid, hier);
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// Returns false if the document is open and there are items on the redo stack
        /// </summary>
        internal bool IsRedoStackEmpty()
        {
            IOleUndoManager? undoManager = UndoManager();
            if (undoManager != null)
            {
                string? description = null;

                // Try casting to internal version of IOleUndoManager
                // that returns hresult instead of throwing
                // this keeps exceptions out of our normal call paths
                if (undoManager is IOleUndoManagerHR undoManagerHR)
                {
                    undoManagerHR.GetLastRedoDescription(out description);
                }
                else
                {
                    try
                    {
                        undoManager.GetLastRedoDescription(out description);
                    }
                    catch
                    {
                    }
                }

                if (!string.IsNullOrEmpty(description))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Gets the undo manager if the document is open
        /// </summary>
        internal IOleUndoManager? UndoManager()
        {
            IVsTextLines? textLines = GetRunningDocumentTextBuffer();
            if (textLines != null)
            {
                textLines.GetUndoManager(out IOleUndoManager undoManager);
                return undoManager;
            }

            return null;
        }

        /// <summary>
        /// Item's id which is unique for the hieararchy
        /// </summary>
        internal uint VsItemID
        {
            get
            {
                return _vsitemid;
            }
        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Read only access to the hierarchy interface
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        internal IVsHierarchy Hierarchy
        //        {
        //            get
        //            {
        //                return _hier;
        //            }
        //        }
        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get the parent VSItemID
        //        /// </summary>
        //        /// <returns>Parent itemid or VSITEMID_NIL</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal uint ParentId
        //        {
        //            get
        //            {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_Parent);

        //                if (o is int)
        //                {
        //                    return unchecked((uint)((int)o));
        //                }
        //                else if (o is uint)
        //                {
        //                    return (uint)o;
        //                }
        //                else
        //                {
        //                    return Microsoft.VisualStudio.VSConstants.VSITEMID_NIL;
        //                }
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get the first child as a VsHierarchyItem
        //        /// </summary>
        //        /// <returns>First child if one exists; null otherwise</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal VsHierarchyItem FirstChild(bool fVisible)
        //        {
        //            uint childId = FirstChildId(fVisible);
        //            if (childId != Microsoft.VisualStudio.VSConstants.VSITEMID_NIL)
        //            {
        //                return new VsHierarchyItem(childId, _hier);
        //            }
        //            return null;
        //        }


        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get the parent as a VsHierarchyItem
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        internal VsHierarchyItem Parent
        //        {
        //            get
        //            {
        //                uint parentId = ParentId;
        //                if (parentId != Microsoft.VisualStudio.VSConstants.VSITEMID_NIL)
        //                {
        //                    return new VsHierarchyItem(parentId, _hier);
        //                }
        //                return null;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get the parent folder VsHierarchyItem. Or null if this is the root node
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public VsHierarchyItem ParentFolder
        //        {
        //            get
        //            {
        //                VsHierarchyItem parent = Parent;
        //                while (parent != null && !parent.IsFolder)
        //                {
        //                    parent = parent.Parent;
        //                }
        //                return parent;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Is the given node an ancestor of this node
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool HasAncestor(VsHierarchyItem item)
        //        {
        //            for (VsHierarchyItem ancestor = Parent; ancestor != null; ancestor = ancestor.Parent)
        //            {
        //                if (ancestor.VsItemID == item.VsItemID)
        //                {
        //                    return true;
        //                }
        //            }
        //            return false;
        //        }
        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns true if this project is a web site or WAP project
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsWebProject
        //        {
        //            get
        //            {
        //                return IsWebSiteProject || IsWebAppProject;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns true if this is a web site project
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsWebSiteProject
        //        {
        //            get
        //            {
        //                Project p = Project;
        //                if (p != null && p.Kind == VsWebSite.PrjKind.prjKindVenusProject)
        //                    return true;
        //                return false;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns true if this is a web application project
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsWebAppProject
        //        {
        //            get
        //            {
        //                return _hier is Microsoft.VisualStudio.Web.Application.IVsWebApplicationProject;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get the root as a VsHierarchyItem
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        internal VsHierarchyItem Root
        //        {
        //            get
        //            {
        //                return new VsHierarchyItem(VSConstants.VSITEMID_ROOT, _hier);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get the id of the next sibling of this item.  
        //        /// </summary>
        //        /// <returns>id of first child if this item has a child.
        //        ///           VSConstants.VSITEMID_NIL is returned otherwise</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal uint NextSiblingId(bool fVisible)
        //        {
        //            object o;
        //            if (fVisible)
        //                o = GetPropHelper(__VSHPROPID.VSHPROPID_NextVisibleSibling);
        //            else
        //                o = GetPropHelper(__VSHPROPID.VSHPROPID_NextSibling);

        //            if (o is int)
        //            {
        //                return unchecked((uint)((int)o));
        //            }
        //            else if (o is uint)
        //            {
        //                return (uint)o;
        //            }
        //            else
        //            {
        //                return Microsoft.VisualStudio.VSConstants.VSITEMID_NIL;
        //            }
        //        }


        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get next sibling for this item.
        //        /// </summary>
        //        /// <returns>VsHierarchItem for next sibling if one exists; null otherwise</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal VsHierarchyItem NextSibling(bool fVisible)
        //        {
        //            uint childId = NextSiblingId(fVisible);
        //            if (childId != Microsoft.VisualStudio.VSConstants.VSITEMID_NIL)
        //            {
        //                return new VsHierarchyItem(childId, _hier);
        //            }
        //            return null;
        //        }
        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Gets the save name for the item.  The save name is the string
        //        /// shown in the save and save changes dialog boxes.
        //        /// </summary>
        //        /// <returns></returns>
        //        //--------------------------------------------------------------------------------------------
        //#if DEADCODE
        //        internal string SaveName
        //        {
        //            get {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_SaveName);
        //                return (o is string) ? (string)o : string.Empty;
        //            }
        //        }
        //#endif

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Gets the string displayed in the project window for a particular item
        //        /// </summary>
        //        /// <returns>Display name for item</returns>
        //        //--------------------------------------------------------------------------------------------
        //#if DEADCODE
        //        internal string Caption
        //        {
        //            get {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_Caption);
        //                return (o is string) ? (string)o : string.Empty;
        //            }
        //        }
        //#endif

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get the name of the item which is basically the file name
        //        /// plus extension minus the directory
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal string Name
        //        {
        //            get
        //            {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_Name);
        //                return (o is string) ? (string)o : string.Empty;
        //            }
        //        }

        /// <summary>
        /// Returns the extensibility object
        /// </summary>
        internal object? ExtObject
        {
            get
            {
                return GetPropHelper(__VSHPROPID.VSHPROPID_ExtObject);
            }
        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the properties IDispatch object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        public object BrowseObject
        //        {
        //            get
        //            {
        //                return GetPropHelper(__VSHPROPID.VSHPROPID_BrowseObject);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the DTE extensibility object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal DTE DTE
        //        {
        //            get
        //            {
        //                DTE dte = null;

        //                VsHierarchyItem root = Root;
        //                if (root != null)
        //                {
        //                    ProjectItems projectItems = root.ProjectItems;
        //                    if (projectItems != null)
        //                    {
        //                        dte = projectItems.DTE;
        //                    }
        //                }

        //                return dte;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the Solution extensibility object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal Solution Solution
        //        {
        //            get
        //            {
        //                Solution solution = null;

        //                DTE dte = DTE;
        //                if (dte != null)
        //                {
        //                    solution = dte.Solution;
        //                }

        //                return solution;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the SolutionBuild extensibility object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal SolutionBuild SolutionBuild
        //        {
        //            get
        //            {
        //                SolutionBuild solutionBuild = null;

        //                Solution solution = Solution;
        //                if (solution != null)
        //                {
        //                    solutionBuild = solution.SolutionBuild;
        //                }

        //                return solutionBuild;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the array of startup projects extensibility object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal string[] StartupProjects()
        //        {
        //            string[] stringStartupProjects = null;

        //            SolutionBuild solutionBuild = SolutionBuild;
        //            if (solutionBuild != null)
        //            {
        //                Array arrayStartupProjects = solutionBuild.StartupProjects as Array;
        //                if (arrayStartupProjects != null && arrayStartupProjects.Length > 0)
        //                {
        //                    stringStartupProjects = new string[arrayStartupProjects.Length];
        //                    arrayStartupProjects.CopyTo(stringStartupProjects, 0);
        //                }
        //            }

        //            return stringStartupProjects;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the true if the hierchy item is in startup project
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal bool IsInStartupProjectList()
        //        {
        //            Project project = Project;
        //            if (project != null)
        //            {
        //                string uniqueName = project.UniqueName;
        //                if (!string.IsNullOrEmpty(uniqueName))
        //                {
        //                    string[] startupProjects = StartupProjects();
        //                    if (startupProjects != null)
        //                    {
        //                        foreach (string name in startupProjects)
        //                        {
        //                            if (name == uniqueName)
        //                            {
        //                                return true;
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            return false;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the Project extensibility object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal Project Project
        //        {
        //            get
        //            {
        //                Project project = null;

        //                VsHierarchyItem root = Root;
        //                if (root != null)
        //                {
        //                    ProjectItems projectItems = root.ProjectItems;
        //                    if (projectItems != null)
        //                    {
        //                        project = projectItems.ContainingProject;
        //                    }
        //                }

        //                return project;
        //            }
        //        }

        /// <summary>
        /// Returns the ProjectItem extensibility object
        /// </summary>
        internal ProjectItem? ProjectItem
        {
            get
            {
                return ExtObject as ProjectItem;
            }
        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the ProjectItems extensibility object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal ProjectItems ProjectItems
        //        {
        //            get
        //            {
        //                if (IsRootNode)
        //                {
        //                    Project project = ExtObject as Project;
        //                    if (project != null)
        //                    {
        //                        return project.ProjectItems;
        //                    }
        //                }
        //                else
        //                {
        //                    ProjectItem projectItem = ProjectItem;
        //                    if (projectItem != null)
        //                    {
        //                        return projectItem.ProjectItems;
        //                    }
        //                }
        //                return null;
        //            }
        //        }
        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the VSWebSite extensibility object for this hierarchy.
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        public VsWebSite.VSWebSite VSWebSite
        //        {
        //            get
        //            {
        //                Project p = Project;
        //                if (p != null)
        //                    return p.Object as VsWebSite.VSWebSite;
        //                return null;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the VsLangProj.References collection from the object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal VSLangProj.VSProject VSProject
        //        {
        //            get
        //            {
        //                if (IsRootNode)
        //                {
        //                    try
        //                    {
        //                        EnvDTE.Project extProj = ExtObject as EnvDTE.Project;
        //                        if (extProj != null)
        //                        {
        //                            VSLangProj.VSProject project = extProj.Object as VSLangProj.VSProject;
        //                            return project;
        //                        }
        //                    }
        //                    catch
        //                    {
        //                    }
        //                }
        //                return null;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns the VsLangProj.References collection from the object
        //        /// </summary>
        //        /// <returns>Name of item</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal VSLangProj.References References
        //        {
        //            get
        //            {
        //                VSLangProj.VSProject project = VSProject;
        //                if (project != null)
        //                {
        //                    return project.References;
        //                }
        //                return null;
        //            }
        //        }
        //        ///------------------------------------------------------------------------
        //        /// <summary>
        //        /// Executes the ExpandItem specified. ie. Expand, collapse, cut highlight, etc
        //        /// </summary>
        //        ///------------------------------------------------------------------------
        //        public void ExpandItem(EXPANDFLAGS flags)
        //        {
        //            IVsUIHierarchyWindow uiWindow = VsShellUtilities.GetUIHierarchyWindow(WAPackage.Package(), new Guid(EnvDTE.Constants.vsWindowKindSolutionExplorer));
        //            if (uiWindow != null)
        //            {
        //                uiWindow.ExpandItem(Hierarchy as IVsUIHierarchy, VsItemID, flags);
        //            }
        //        }

        /// <summary>
        /// Returns the the full path of the item using extensibility.
        /// </summary>
        internal string FullPath
        {
            get
            {
                object? objValue = GetProperty("FullPath");
                if (objValue != null && objValue is string strVal)
                {
                    return strVal;
                }

                return string.Empty;
            }
        }

        //        ///--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Returns the relative path of the project item
        //        ///
        //        ///     folder\file.ext
        //        /// </summary>
        //        ///--------------------------------------------------------------------------------------------
        //        internal string ProjRelativePath
        //        {
        //            get
        //            {
        //                string projRelativePath = null;

        //                string rootProjectDir = Root.ProjectDir;
        //                rootProjectDir = WAUtil.EnsureTrailingBackSlash(rootProjectDir);
        //                string fullPath = FullPath;

        //                if (!string.IsNullOrEmpty(rootProjectDir) && !string.IsNullOrEmpty(fullPath))
        //                {
        //                    projRelativePath = WAUtil.MakeRelativePath(fullPath, rootProjectDir);
        //                }

        //                return projRelativePath;
        //            }
        //        }

        //        ///--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Returns the relative path of the project item
        //        ///
        //        ///     folder/file.ext
        //        /// </summary>
        //        ///--------------------------------------------------------------------------------------------
        //        internal string ProjRelativeUrl
        //        {
        //            get
        //            {
        //                string projRelativeUrl = null;

        //                string projRelativePath = ProjRelativePath;
        //                if (!string.IsNullOrEmpty(projRelativePath))
        //                {
        //                    projRelativeUrl = projRelativePath.Replace(Path.DirectorySeparatorChar, '/');
        //                }

        //                return projRelativeUrl;
        //            }
        //        }

        //        ///--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Returns the application root directory (which can be differenct than project dir)
        //        ///
        //        ///     c:\app
        //        /// </summary>
        //        ///--------------------------------------------------------------------------------------------
        //        internal string AppDir
        //        {
        //            get
        //            {
        //                string appDir = null;

        //                WAProject waProject = WAProject.GetProjectFromIVsHierarchy(_hier);
        //                if (waProject != null)
        //                {
        //                    appDir = waProject.AppDirectory;
        //                }

        //                return appDir;
        //            }
        //        }

        //        ///--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Returns the relative path of the project item to the app
        //        ///
        //        ///     folder\file.ext
        //        /// </summary>
        //        ///--------------------------------------------------------------------------------------------
        //        internal string AppRelativePath
        //        {
        //            get
        //            {
        //                string appRelativePath = null;

        //                string rootAppDir = AppDir;
        //                string fullPath = FullPath;

        //                if (!string.IsNullOrEmpty(rootAppDir) && !string.IsNullOrEmpty(fullPath))
        //                {
        //                    appRelativePath = WAUtil.MakeRelativePath(fullPath, rootAppDir);
        //                }

        //                return appRelativePath;
        //            }
        //        }

        //        ///--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Returns the app relative Url of the project item
        //        ///
        //        ///     folder/file.ext
        //        /// </summary>
        //        ///--------------------------------------------------------------------------------------------
        //        internal string AppRelativeUrl
        //        {
        //            get
        //            {
        //                string appRelativeUrl = null;

        //                string appRelativePath = AppRelativePath;
        //                if (!string.IsNullOrEmpty(appRelativePath))
        //                {
        //                    appRelativeUrl = appRelativePath.Replace(Path.DirectorySeparatorChar, '/');
        //                }

        //                return appRelativeUrl;
        //            }
        //        }

        //        ///--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Returns the app relative Url of the project item
        //        ///     with the ASP ~/ prefix
        //        ///
        //        ///     ~/folder/file.ext
        //        /// </summary>
        //        ///--------------------------------------------------------------------------------------------
        //        internal string AspAppRelativeUrl
        //        {
        //            get
        //            {
        //                string aspAppRelativeUrl = null;

        //                string appRelativeUrl = AppRelativeUrl;
        //                if (!string.IsNullOrEmpty(appRelativeUrl))
        //                {
        //                    aspAppRelativeUrl = "~/" + appRelativeUrl;
        //                }

        //                return aspAppRelativeUrl;
        //            }
        //        }

        //        ///--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Returns the full Url to browse to this item unter the current web server
        //        ///
        //        ///     http://localhost:8021/folder/file.ext
        //        /// </summary>
        //        ///--------------------------------------------------------------------------------------------
        //        internal string BrowseUrl
        //        {
        //            get
        //            {
        //                string browseUrl = null;

        //                WAProject waProject = WAProject.GetProjectFromIVsHierarchy(_hier);
        //                if (waProject != null)
        //                {
        //                    string serverUrl = waProject.BrowseUrlTS;
        //                    string projRelativeUrl = ProjRelativeUrl;
        //                    if (!string.IsNullOrEmpty(projRelativeUrl) && !string.IsNullOrEmpty(serverUrl))
        //                    {
        //                        browseUrl = serverUrl + projRelativeUrl;
        //                    }
        //                }

        //                return browseUrl;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is web.config
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsWebConfig
        //        {
        //            get
        //            {
        //                if (IsFile)
        //                {
        //                    string name = Name;
        //                    if (!string.IsNullOrEmpty(name))
        //                    {
        //                        const string webConfig = "web.config";
        //                        if (name.Length == webConfig.Length
        //                            && string.Compare(name, webConfig, StringComparison.OrdinalIgnoreCase) == 0)
        //                        {
        //                            return true;
        //                        }
        //                    }
        //                }
        //                return false;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is App_Code
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //#if DEADCODE
        //        public bool IsAppCode
        //        {
        //            get {
        //                return IsSpecialFolder(SpecialFolder.AppCode);
        //            }
        //        }
        //#endif

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is App_Themes
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //#if DEADCODE
        //        public bool IsAppThemes
        //        {
        //            get {
        //                return IsSpecialFolder(SpecialFolder.AppThemes);
        //            }
        //        }
        //#endif

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is App_Browsers
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //#if DEADCODE
        //        public bool IsAppBrowsers
        //        {
        //            get {
        //                return IsSpecialFolder(SpecialFolder.AppBrowsers);
        //            }
        //        }
        //#endif

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is the requsted special directory
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsSpecialFolder(SpecialFolder specialFolder)
        //        {
        //            if (specialFolder != null && IsFolder && (!specialFolder.MustBeAtRoot || ParentId == VSConstants.VSITEMID_ROOT))
        //            {
        //                string name = Name;
        //                string specialFolderName = specialFolder.Name;
        //                if (!string.IsNullOrEmpty(name)
        //                    && !string.IsNullOrEmpty(specialFolderName)
        //                    && name.Length == specialFolderName.Length
        //                    && string.Compare(name, specialFolderName, StringComparison.OrdinalIgnoreCase) == 0)
        //                {
        //                    return true;
        //                }
        //            }
        //            return false;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is excluded from the project
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsExcludedFromProject
        //        {
        //            get
        //            {
        //                object val = GetPropHelper(__VSHPROPID.VSHPROPID_IsNonMemberItem);
        //                return (val is bool) && (bool)val;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is under App_Code or is App_Code
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderAppCode
        //        {
        //            get
        //            {
        //                return IsUnderSpecialFolder(SpecialFolder.AppCode);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is under App_Data or is App_Data
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderAppData
        //        {
        //            get
        //            {
        //                return IsUnderSpecialFolder(SpecialFolder.AppData);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is under App_Themes or is App_Themes
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderAppThemes
        //        {
        //            get
        //            {
        //                return IsUnderSpecialFolder(SpecialFolder.AppThemes);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is under App_Browsers or is App_Browsers
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderAppBrowsers
        //        {
        //            get
        //            {
        //                return IsUnderSpecialFolder(SpecialFolder.AppBrowsers);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is under App_GlobalResources or is App_GlobalResources
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderAppGlobalResources
        //        {
        //            get
        //            {
        //                return IsUnderSpecialFolder(SpecialFolder.AppGlobalResources);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is under App_LocalResources or is App_LocalResources
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderAppLocalResources
        //        {
        //            get
        //            {
        //                return IsUnderSpecialFolder(SpecialFolder.AppLocalResources);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns true if this item is under the requested special
        //        ///     directory or is the special directory
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderSpecialFolder(SpecialFolder specialFolder)
        //        {
        //            for (VsHierarchyItem item = this; item != null; item = item.Parent)
        //            {
        //                if (item.IsSpecialFolder(specialFolder))
        //                {
        //                    return true;
        //                }
        //            }
        //            return false;
        //        }


        //        /// <summary>
        //        /// Return true if this item is "Service References" or under it
        //        /// </summary>
        //        /// <returns></returns>
        //        public bool IsUnderServiceReference()
        //        {
        //            VsHierarchyItem topLevelFolder = GetTopFolderContainingItem();
        //            if (topLevelFolder != null)
        //            {
        //                string name = topLevelFolder.Name;
        //                if (string.Equals(SpecialFolder.ServiceReferences.Name, name, StringComparison.OrdinalIgnoreCase))
        //                {
        //                    return true;
        //                }
        //            }
        //            return false;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns true if this item is under any of the app_* folders.
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderASpecialFolder()
        //        {
        //            VsHierarchyItem topLevelFolder = GetTopFolderContainingItem();
        //            if(topLevelFolder != null)
        //            {
        //                string name = topLevelFolder.Name;
        //                if(name.StartsWith("App_", StringComparison.OrdinalIgnoreCase) &&
        //                  (SpecialFolder.AppCode.Name.Equals(name, StringComparison.OrdinalIgnoreCase) ||
        //                   SpecialFolder.AppData.Name.Equals(name, StringComparison.OrdinalIgnoreCase) ||
        //                   SpecialFolder.AppGlobalResources.Name.Equals(name, StringComparison.OrdinalIgnoreCase) ||
        //                   SpecialFolder.AppThemes.Name.Equals(name, StringComparison.OrdinalIgnoreCase) ||
        //                   SpecialFolder.AppBrowsers.Name.Equals(name, StringComparison.OrdinalIgnoreCase) ||
        //                   SpecialFolder.AppWebReferences.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
        //                {
        //                    return true;
        //                }
        //            }
        //            // Special case app local resources since it can be at any level
        //            return  IsUnderAppLocalResources;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///  Returns the top level folder containing this item, or this item if it happens to be
        //        ///  a top level folder. Files at the root return NULL.
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        internal VsHierarchyItem GetTopFolderContainingItem()
        //        {
        //            for (VsHierarchyItem item = this; item != null; item = item.Parent)
        //            {
        //                if (item.IsFolder && item.ParentId == VSConstants.VSITEMID_ROOT)
        //                {
        //                    return item;
        //                }
        //            }
        //            return null;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///  returns true if this item is under the requested folder name or is the matching folder
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public bool IsUnderFolder(string folderName)
        //        {
        //            for (VsHierarchyItem item = this; item != null; item = item.Parent)
        //            {
        //                if (item.IsFolder && string.Equals(Name, folderName, StringComparison.OrdinalIgnoreCase))
        //                {
        //                    return true;
        //                }
        //            }
        //            return false;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Return the TFM. Always goes off the root
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public string TargetFrameworkMoniker
        //        {
        //            get
        //            {
        //                object objMoniker = GetPropHelper((uint)VSConstants.VSITEMID_ROOT, (int)__VSHPROPID4.VSHPROPID_TargetFrameworkMoniker);
        //                if (objMoniker is string)
        //                    return (string)objMoniker;
        //                return "";
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     returns the specified folder if exists
        //        ///     if the folder does not exist it creates and returns a new one.
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public VsHierarchyItem EnsureChildFolder(string folderName)
        //        {
        //            VsHierarchyItem child = null;

        //            if (!string.IsNullOrEmpty(folderName) && IsFolder)
        //            {
        //                child = GetChildFolder(folderName);
        //                if (child == null)
        //                {
        //                    child = CreateChildFolder(folderName);
        //                }
        //            }
        //            return child;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     creates the specified folder
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public VsHierarchyItem CreateChildFolder(string folderName)
        //        {
        //            if (!string.IsNullOrEmpty(folderName) && IsFolder)
        //            {
        //                ProjectItems children = ProjectItems;
        //                if (children != null)
        //                {
        //                    children.AddFolder(folderName, null);
        //                    return GetChildFolder(folderName);
        //                }
        //            }
        //            return null;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     gets the child folder by name
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public VsHierarchyItem GetChildFolder(string folderName)
        //        {
        //            if (!string.IsNullOrEmpty(folderName) && IsFolder)
        //            {
        //                bool visible = false;
        //                VsHierarchyItem current = FirstChild(visible);
        //                while (current != null)
        //                {
        //                    if (current.IsFolder)
        //                    {
        //                        string name = current.Name;
        //                        if (!string.IsNullOrEmpty(name)
        //                            && name.Length == folderName.Length
        //                            && string.Compare(name, folderName, StringComparison.OrdinalIgnoreCase) == 0)
        //                        {
        //                            return current;
        //                        }
        //                    }
        //                    current = current.NextSibling(visible);
        //                }
        //            }
        //            return null;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     gets the child folder by name
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public VsHierarchyItem GetChildOfName(string itemName)
        //        {
        //            VsHierarchyItem child = FirstChild(false /*fVisible*/);
        //            while (child != null)
        //            {
        //                // Look for matching item
        //                if (string.Compare(child.Name, itemName, StringComparison.OrdinalIgnoreCase) == 0)
        //                {
        //                    return child;
        //                }
        //                child = child.NextSibling(false /*fVisible*/);
        //            }
        //            return null;
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns whether this item is a container for other items
        //        /// </summary>
        //        /// <returns>true if this is a container, false otherwise</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal bool IsExpandable
        //        {
        //            get
        //            {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_Expandable);

        //                if (o is bool)
        //                {
        //                    return (bool)o;
        //                }
        //                return (o is int) ? (int)o != 0 : false;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns whether this item is a link file
        //        /// </summary>
        //        /// <returns>true if this is a container, false otherwise</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal bool IsLinkFile
        //        {
        //            get
        //            {
        //                object o = GetPropHelper(__VSHPROPID2.VSHPROPID_IsLinkFile);

        //                if (o is bool)
        //                {
        //                    return (bool)o;
        //                }
        //                return (o is int) ? (int)o != 0 : false;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns whether this item is hidden or not
        //        /// </summary>
        //        /// <returns>true if hidden, false otherwise</returns>
        //        //--------------------------------------------------------------------------------------------
        //#if DEADCODE
        //        internal bool IsHidden
        //        {
        //            get {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_IsHiddenItem);

        //                if(o is bool)
        //                {
        //                    return (bool)o;
        //                }
        //                return (o is int) ? (int)o != 0 : false;
        //            }
        //        }
        //#endif

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns whether this item is a member or not
        //        /// </summary>
        //        /// <returns>true if hidden, false otherwise</returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal bool IsMemberItem
        //        {
        //            get
        //            {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_IsNonMemberItem);

        //                //  If property isn't supported or an error, assume it is a member
        //                if (o == null || !(o is bool))
        //                    return true;
        //                else
        //                    return !(bool)o;    // not nonmembers are members.
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Gets full path to project directory for this item
        //        /// </summary>
        //        /// <returns></returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal string ProjectDir
        //        {
        //            get
        //            {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_ProjectDir);
        //                return (o is string) ? (string)o : string.Empty;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get type name for this item.  This is the display name used in the 
        //        /// title bar to identify the type of the node or hierarchy.
        //        /// </summary>
        //        /// <returns></returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal bool IsRootNode
        //        {
        //            get
        //            {
        //                return _vsitemid == VSConstants.VSITEMID_ROOT;
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Get type name for this item.  This is the display name used in the 
        //        /// title bar to identify the type of the node or hierarchy.
        //        /// </summary>
        //        /// <returns></returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal bool IsReferencesNode
        //        {
        //            get
        //            {
        //                if (_isReferencesVFolder == -1)
        //                {
        //                    _isReferencesVFolder = 0;
        //                    if (ParentId == VSConstants.VSITEMID_ROOT)
        //                    {
        //                        string name = Name;
        //                        const string specialFolderName = "References";
        //                        if (!string.IsNullOrEmpty(name)
        //                            && !string.IsNullOrEmpty(specialFolderName)
        //                            && name.Length == specialFolderName.Length
        //                            && string.Compare(name, specialFolderName, StringComparison.OrdinalIgnoreCase) == 0)
        //                        {
        //                            Guid guid = TypeGuid;
        //                            if (guid.CompareTo(VSConstants.GUID_ItemType_VirtualFolder) == 0)
        //                            {
        //                                _isReferencesVFolder = 1;
        //                            }
        //                        }
        //                    }
        //                }
        //                return _isReferencesVFolder == 1;
        //            }
        //        }

        /// <summary>
        /// Gets the type guid for this item and compares against GUID_ItemType_PhysicalFile
        /// </summary>
        /// <returns></returns>
        internal bool IsFile
        {
            get
            {
                if (_isFile == -1)
                {
                    Guid guid = TypeGuid;
                    _isFile = guid.CompareTo(VSConstants.GUID_ItemType_PhysicalFile) == 0 ? 1 : 0;
                }
                return _isFile == 1;
            }
        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Gets the type guid for this item and compares against GUID_ItemType_PhysicalFolder. Note
        //        /// that the root node is considered a folder.
        //        /// </summary>
        //        /// <returns></returns>
        //        //--------------------------------------------------------------------------------------------
        //        internal bool IsFolder
        //        {
        //            get
        //            {
        //                if (_isFolder == -1)
        //                {
        //                    if (IsRootNode)
        //                        _isFolder = 1;
        //                    else
        //                    {
        //                        Guid guid = TypeGuid;
        //                        _isFolder = guid.CompareTo(VSConstants.GUID_ItemType_PhysicalFolder) == 0 ? 1 : 0;
        //                    }
        //                }
        //                return _isFolder == 1;
        //            }
        //        }

        /// <summary>
        /// Indicates if the file is opened in a document editor and modified.
        /// </summary>
        /// <returns></returns>
        public bool IsDirty
        {
            get
            {
                ProjectItem ?projectItem = ProjectItem;
                if (projectItem != null)
                {
                    Document? document = projectItem.Document;
                    if (document != null && !document.Saved)
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        //        ///-------------------------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     If the item is open in a document editor it returns the window frame
        //        /// </summary>
        //        ///-------------------------------------------------------------------------------------------------------------
        //        public IVsWindowFrame DocumentFrame(uint grfIDO = 0)
        //        {
        //            IVsWindowFrame frame = null;

        //            IVsUIHierarchy uiHierarchy = _hier as IVsUIHierarchy;
        //            if (uiHierarchy != null && IsFile)
        //            {
        //                IVsUIShellOpenDocument shell = GetService<IVsUIShellOpenDocument>();
        //                if (shell != null)
        //                {
        //                    Guid logicalView = Guid.Empty;
        //                    IVsUIHierarchy uiHierarchyOpen;
        //                    uint[] itemidOpen = new uint[1];
        //                    int isOpen;
        //                    shell.IsDocumentOpen(uiHierarchy, _vsitemid, FullPath, ref logicalView, grfIDO, out uiHierarchyOpen, itemidOpen, out frame, out isOpen);
        //                }
        //            }
        //            return frame;
        //        }

        //        /*
        //                ///-------------------------------------------------------------------------------------------------------------
        //                /// <summary>
        //                ///     If the item is open in a document editor it returns the editor view
        //                /// </summary>
        //                ///-------------------------------------------------------------------------------------------------------------
        //                public object DocumentView()
        //                {
        //                    object view = null;

        //                    IVsWindowFrame frame = DocumentFrame();
        //                    if (frame != null)
        //                    {
        //                        frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out view);
        //                    }

        //                    return view;
        //                }

        //                ///-------------------------------------------------------------------------------------------------------------
        //                /// <summary>
        //                ///     If the item is open in a document editor it returns the editor view command target
        //                /// </summary>
        //                ///-------------------------------------------------------------------------------------------------------------
        //                public IOleCommandTarget DocumentViewCommandTarget()
        //                {
        //                    IOleCommandTarget viewCmdTarget = null;

        //                    object view = DocumentView();
        //                    if (view != null)
        //                    {
        //                        viewCmdTarget = view as IOleCommandTarget;
        //                    }

        //                    return viewCmdTarget;
        //                }
        //        */

        //#if DEADCODE
        //        ///-------------------------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     If the file is already open it shows it in the current view
        //        ///     If the file is not open it opens it in the default view
        //        /// </summary>
        //        ///-------------------------------------------------------------------------------------------------------------
        //        public IVsWindowFrame Show()
        //        {
        //            IVsWindowFrame frame = DocumentFrame();
        //            if (frame != null)
        //            {
        //                frame.Show();
        //                return frame;
        //            }
        //            return Open();
        //        }

        //        ///-------------------------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Open the file in default view
        //        /// </summary>
        //        ///-------------------------------------------------------------------------------------------------------------
        //        public IVsWindowFrame Open()
        //        {
        //            return Open(VSConstants.LOGVIEWID_Primary);
        //        }
        //#endif
        //        ///-------------------------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Open the file in design view
        //        /// </summary>
        //        ///-------------------------------------------------------------------------------------------------------------
        //        public IVsWindowFrame OpenDesigner()
        //        {
        //            return Open(VSConstants.LOGVIEWID_Designer);
        //        }

        //        ///-------------------------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Open the file in code view
        //        /// </summary>
        //        ///-------------------------------------------------------------------------------------------------------------
        //        public IVsWindowFrame OpenCode()
        //        {
        //            return Open(VSConstants.LOGVIEWID_Code);
        //        }

        /// <summary>
        /// Returns the current contents of a document
        /// </summary>
        public string GetDocumentText()
        {
            string text = string.Empty;
            IVsPersistDocData? docData = null;

            try
            {
                // Get or create the buffer
                IVsTextLines? buffer = GetRunningDocumentTextBuffer();
                if (buffer == null)
                {
                    docData = CreateDocumentData();
                    buffer = docData as IVsTextLines;
                }

                // get the text from the buffer
                if (buffer != null)
                {
                    if (buffer is IVsTextStream textStream)
                    {
                        int hr = textStream.GetSize(out int length);
                        if (ErrorHandler.Succeeded(hr))
                        {
                            if (length > 0)
                            {
                                IntPtr pText = Marshal.AllocCoTaskMem((length + 1) * 2);
                                try
                                {
                                    hr = textStream.GetStream(0, length, pText);
                                    if (ErrorHandler.Succeeded(hr))
                                    {
                                        text = Marshal.PtrToStringUni(pText);
                                    }
                                }
                                finally
                                {
                                    Marshal.FreeCoTaskMem(pText);
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                if (docData != null)
                {
                    docData.Close();
                }
            }

            return text;
        }

        /// <summary>
        /// Creates and loads the document data
        /// (You must Close() it when done)
        /// </summary>
        public IVsPersistDocData? CreateDocumentData()
        {
            if (IsFile)
            {
                string fullPath = FullPath;
                if (!string.IsNullOrEmpty(fullPath))
                {
                    IOleServiceProvider? site = Site();
                    if (site != null)
                    {
                        IVsPersistDocData? docData = CodeGenUtils.CreateSitedInstance<IVsPersistDocData>(site, typeof(VsTextBufferClass).GUID);
                        if (docData != null)
                        {
                            int hr = docData.LoadDocData(fullPath);
                            if (ErrorHandler.Succeeded(hr))
                            {
                                return docData;
                            }
                            else
                            {   // We need to close the docdata when we are done with it.
                                docData.Close();
                            }
                        }
                    }
                }
            }
            return null;
        }

        //        ///-------------------------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// If the document is open, it returns the IVsWebApplicationDocument.
        //        /// </summary>
        //        ///-------------------------------------------------------------------------------------------------------------
        //        public IVsWebApplicationDocument GetRunningWebApplicationDocument()
        //        {
        //            return GetRunningDocumentData() as IVsWebApplicationDocument;
        //        }

        /// <summary>
        /// If the document is open, it returns the IVsTextLines.
        /// </summary>
        public IVsTextLines? GetRunningDocumentTextBuffer()
        {
            IVsTextLines? buffer = null;

            IVsPersistDocData? docData = GetRunningDocumentData();
            if (docData != null)
            {
                buffer = docData as IVsTextLines;
                if (buffer == null)
                {
                    if (docData is IVsTextBufferProvider provider)
                    {
                        provider.GetTextBuffer(out buffer);
                    }
                }
            }

            return buffer;
        }

        /// <summary>
        /// If the document is open, it returns the IVsPersistDocData for it.
        /// </summary>
        public IVsPersistDocData? GetRunningDocumentData()
        {
            IVsPersistDocData? persistDocData = null;

            IntPtr docData = IntPtr.Zero;
            try
            {
                docData = GetRunningDocData();
                if (docData != IntPtr.Zero)
                {
                    persistDocData = Marshal.GetObjectForIUnknown(docData) as IVsPersistDocData;
                }
            }
            finally
            {
                if (docData != IntPtr.Zero)
                {
                    Marshal.Release(docData);
                }
            }

            return persistDocData;
        }

        /// <summary>
        /// If the document is open, it returns the IntPtr to the doc data.
        /// (This is ref-counted and must be released with Marshal.Release())
        /// </summary>
        public IntPtr GetRunningDocData()
        {
            IntPtr docData = IntPtr.Zero;

            if (IsFile)
            {
                string? fullPath = FullPath;
                if (!Strings.IsNullOrEmpty(fullPath))
                {
                    IVsRunningDocumentTable? rdt = GetService<IVsRunningDocumentTable>();
                    if (rdt != null)
                    {
                        _VSRDTFLAGS flags = _VSRDTFLAGS.RDT_NoLock;
                        rdt.FindAndLockDocument
                        (
                            (uint)flags,
                            fullPath,
                            out _,
                            out uint _,
                            out docData,
                            out uint _
                        );
                    }
                }
            }

            return docData;
        }

        //        ///-------------------------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Open in specified logical view
        //        /// </summary>
        //        ///-------------------------------------------------------------------------------------------------------------
        //        public IVsWindowFrame Open(Guid logicalView)
        //        {
        //            IVsWindowFrame windowFrame = null;

        //            if (_hier != null && IsFile)
        //            {
        //                IntPtr docData = IntPtr.Zero;
        //                string fullPath = FullPath;

        //                try
        //                {
        //                    docData = GetRunningDocData();

        //                    IVsUIShellOpenDocument uiShellOpenDocument = GetService<IVsUIShellOpenDocument>();
        //                    if (uiShellOpenDocument != null)
        //                    {
        //                        __VSOSEFLAGS openFlags = __VSOSEFLAGS.OSE_OpenAsNewFile;
        //                        openFlags |= __VSOSEFLAGS.OSE_ChooseBestStdEditor;
        //                        IOleServiceProvider serviceProvider = Site();
        //                        IVsUIHierarchy uiHierarchy = UIHierarchy;

        //                        if (serviceProvider != null && uiHierarchy != null)
        //                        {
        //                            uiShellOpenDocument.OpenStandardEditor
        //                            (
        //                               (uint)openFlags,
        //                               fullPath,
        //                               ref logicalView,
        //                               "%2",
        //                               uiHierarchy,
        //                               _vsitemid,
        //                               docData,
        //                               serviceProvider,
        //                               out windowFrame
        //                            );
        //                        }
        //                    }
        //                }
        //                finally
        //                {
        //                    if (docData != IntPtr.Zero)
        //                    {
        //                        Marshal.Release(docData);
        //                    }
        //                }

        //                if (windowFrame != null)
        //                {
        //                    windowFrame.Show();
        //                }
        //            }

        //            return windowFrame;
        //        }

        /// <summary>
        ///     Gets the project item Properties collection.
        /// </summary>
        public EnvDTE.Properties? Properties
        {
            get
            {
                object? obj = ExtObject;
                if (obj is EnvDTE.Project project)
                {
                    return project.Properties;
                }
                else if (obj is EnvDTE.ProjectItem item)
                {
                    return item.Properties;
                }

                return null;
            }
        }

        /// <summary>
        /// Helper to get the named property. Returns null if not found
        /// </summary>
        public object? GetProperty(string propName)
        {
            EnvDTE.Properties? props = Properties;
            if (props != null)
            {
                // Unfortunately, the code throws if the property doesn't exist...
                try
                {
                    Property? property = props.Item(propName);
                    if (property != null)
                    {
                        return property.Value;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        /// <summary>
        ///     Gets the build action of the item.
        ///     If not found returns build action none.
        /// </summary>
        public prjBuildAction BuildAction
        {
            get
            {
                object? objValue = GetProperty("BuildAction");
                if (objValue != null)
                {
                    prjBuildAction buildAction = (prjBuildAction)objValue;
                    return buildAction;
                }

                return prjBuildAction.prjBuildActionNone;
            }
        }

        //        ///-------------------------------------------------------------------------------------------------------------
        //        /// <summary>
        //        ///     Gets the custom tool of the item.
        //        /// </summary>
        //        ///-------------------------------------------------------------------------------------------------------------
        //#if DEADCODE
        //        public string CustomTool()
        //        {
        //            get {
        //                object objValue = GetProperty("CustomTool");
        //                if (objValue != null)
        //                {
        //                    return objValue as string;
        //                }
        //                return null;
        //            }
        //        }
        //#endif

        /// <summary>
        ///     Returns true if the build action of the item is compile.
        /// </summary>
        public bool IsBuildActionCompile
        {
            get
            {
                prjBuildAction buildAction = BuildAction;
                if (buildAction == prjBuildAction.prjBuildActionCompile)
                {
                    return true;
                }
                return false;
            }
        }

        /// <summary>
        /// Gets the default namespace for the item.
        /// </summary>
        /// <returns></returns>
        public string? DefaultNamespace
        {
            get
            {
                object? obj = GetPropHelper(__VSHPROPID.VSHPROPID_DefaultNamespace);
                return (obj as string);
            }
        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Gets the state icon for the hiearchy item.
        //        /// typedef enum __VSSTATEICON {
        //        ///   STATEICON_NOSTATEICON              = 0,
        //        ///   STATEICON_CHECKEDIN                = 1,
        //        ///   STATEICON_CHECKEDOUT               = 2,
        //        ///   STATEICON_ORPHANED                 = 3,
        //        ///   STATEICON_EDITABLE                 = 4,
        //        ///   STATEICON_BLANK                    = 5,
        //        ///   STATEICON_READONLY                 = 6,
        //        ///   STATEICON_DISABLED                 = 7,
        //        ///   STATEICON_CHECKEDOUTEXCLUSIVE      = 8,
        //        ///   STATEICON_CHECKEDOUTSHAREDOTHER    = 9,
        //        ///   STATEICON_CHECKEDOUTEXCLUSIVEOTHER = 10,
        //        ///   STATEICON_EXCLUDEDFROMSCC          = 11,
        //        ///   STATEICON_MAXINDEX                 = 12
        //        /// } VsStateIcon;
        //        /// </summary>
        //        /// <param name="itemid"></param>
        //        /// <returns></returns>
        //        //--------------------------------------------------------------------------------------------
        //#if DEADCODE
        //        internal VsStateIcon VssState
        //        {
        //            get {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_StateIconIndex);
        //                return (o is int) ? (VsStateIcon)((int)o) : VsStateIcon.STATEICON_BLANK;
        //            }
        //        }
        //#endif

        //#if DEADCODE
        //        internal object IconHandle()
        //        {
        //            get {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_IconHandle);
        //                return o;
        //            }
        //        }
        //#endif

        //#if DEADCODE
        //        internal object IconImageList
        //        {
        //            get {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_IconImgList);
        //                return o;
        //            }
        //        }
        //#endif

        //#if DEADCODE
        //        internal int GIconIndex()
        //        {
        //            get {
        //                object o = GetPropHelper(__VSHPROPID.VSHPROPID_IconIndex);
        //                return (o is int) ? (int)o : -1;
        //            }
        //        }
        //#endif


        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns AggregateProjectTypeGuids
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        public string AggregateProjectTypeGuids
        //        {
        //            get
        //            {
        //                string projectTypeGuids = null;
        //                IVsAggregatableProjectCorrected aggregate = _hier as IVsAggregatableProjectCorrected;
        //                if (aggregate != null)
        //                {
        //                    aggregate.GetAggregateProjectTypeGuids(out projectTypeGuids);
        //                }
        //                return projectTypeGuids;
        //            }
        //        }

        /// <summary>
        /// Returns VSHPROPID_TypeGuid
        /// </summary>
        internal Guid TypeGuid
        {
            get
            {
                return GetGuidPropHelper(_vsitemid, (int)__VSHPROPID.VSHPROPID_TypeGuid);
            }
        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returnss VSHPROPID_ProjectIDGuid
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        internal Guid ProjectIdGuid
        //        {
        //            get
        //            {
        //                return GetGuidPropHelper(_vsitemid, (int)__VSHPROPID.VSHPROPID_ProjectIDGuid);
        //            }
        //        }

        //        //--------------------------------------------------------------------------------------------
        //        /// <summary>
        //        /// Returns VSHPROPID_AddItemTemplatesGuid
        //        /// </summary>
        //        //--------------------------------------------------------------------------------------------
        //        internal Guid AddItemTemplatesGuid
        //        {
        //            get
        //            {
        //                //  Always from the root
        //                return GetGuidPropHelper(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID2.VSHPROPID_AddItemTemplatesGuid);
        //            }
        //        }

        /// <summary>
        /// Get the specified property from the __VSHPROPID enumeration for this item
        /// </summary>
        private object? GetPropHelper(__VSHPROPID propid)
        {
            return GetPropHelper(_vsitemid, (int)propid);
        }

        /// <summary>
        /// Get the specified property from the __VSHPROPID2 enumeration for this item
        /// </summary>
        private object? GetPropHelper(__VSHPROPID2 propid)
        {
            return GetPropHelper(_vsitemid, (int)propid);
        }

        /// <summary>
        /// Get the specified property for the specified item
        /// </summary>
        private object? GetPropHelper(uint itemid, int propid)
        {
            try
            {
                object? o = null;

                if (_hier != null)
                {
                    int hr = _hier.GetProperty(itemid, propid, out o);
                }

                return o;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Calls IVsHIerachy::GetGuidProperty
        /// </summary>
        internal Guid GetGuidPropHelper(uint itemid, int propid)
        {
            Guid guid;
            try
            {
                _hier.GetGuidProperty(itemid, propid, out guid);
            }
            catch (Exception)
            {
                guid = Guid.Empty;
            }

            return guid;
        }

        /// <summary>
        /// Get hierarchy site
        /// </summary>
        public IOleServiceProvider? Site()
        {
            IOleServiceProvider? serviceProvider = null;
            if (_hier != null)
            {
                _hier.GetSite(out serviceProvider);
            }

            return serviceProvider;
        }

        /// <summary>
        /// Helper to get a shell service interface
        /// </summary>
        public InterfaceType? GetService<InterfaceType>() where InterfaceType : class
        {
            InterfaceType? service = null;

            try
            {
                if (_serviceProvider == null)
                {
                    IOleServiceProvider? serviceProvider = Site();
                    if (serviceProvider != null)
                    {
                        _serviceProvider = new ServiceProvider(serviceProvider);
                    }
                }

                if (_serviceProvider != null)
                {
#pragma warning disable RS0030 // Do not used banned APIs
                    service = _serviceProvider.GetService(typeof(InterfaceType)) as InterfaceType;
#pragma warning restore RS0030 // Do not used banned APIs
                }
            }
            catch
            {
            }

            return service;
        }
    }
}

