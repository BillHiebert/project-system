using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Runtime.Versioning;
using System.Web;
using System.Web.Configuration;
using System.Web.UI;
using System.Web.UI.Design;
using System.Web.UI.HtmlControls;
using Microsoft.VisualStudio.Shell.Interop;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    internal class Parser
    {
        // Construction state
        private readonly string _appVirtualPath;               // The virtual path of the app root
        private readonly string _appPhysicalPath;              // The physical directory of the app root
        private readonly IDesignerHost _designerHost;                 // Host for type resolving
        private readonly IVsHierarchy _hierarchy;

        // BeginParse state
        private string? _virtualPath;             // The file to be parsed
        private string? _text;                    // The file contents

        // ParseDirectives state
        private string? _className_Full;
        private string? _className_Name;
        private string? _className_Namespace;
        private string? _masterPageTypeName;
        private string? _previousPageTypeName;
        private UserControlTypeNameDictionary? _userControls;

        // ParseConfig state
        private DirectiveList? _configRegisterDirectives;

        // Parse state
        private DirectiveList? _registerDirectives;
        private Hashtable? _htmlTagMap;
        private Hashtable? _htmlInputTypes;
        private ControlInfoList? _controlInfos;
        private Version? _targetFrameworkVersion;

        // Static versions
        private static readonly Version s_framework45 = new Version(4, 5);
        private static readonly Version s_latestFramework = CurrentFrameworkVersion();

        public Parser(string appVirtualPath, string appPhysicalPath, IDesignerHost designerHost, IVsHierarchy hierarchy)
        {
            if (appVirtualPath == null)
            {
                throw new ArgumentNullException(nameof(appVirtualPath));
            }

            if (appPhysicalPath == null)
            {
                throw new ArgumentNullException(nameof(appPhysicalPath));
            }

            if (designerHost == null)
            {
                throw new ArgumentNullException(nameof(designerHost));
            }

            if (hierarchy == null)
            {
                throw new ArgumentNullException(nameof(hierarchy));
            }

            // Make sure the app's virtual and physical paths end with a slash
            appVirtualPath = VirtualPathUtility.AppendTrailingSlash(appVirtualPath);
            if (appPhysicalPath[appPhysicalPath.Length - 1] != '\\')
            {
                appPhysicalPath += '\\';
            }

            _appVirtualPath = appVirtualPath;
            _appPhysicalPath = appPhysicalPath;
            _designerHost = designerHost;
            _hierarchy = hierarchy;
        }

        public string? ClassName_Full
        {
            get
            {
                return _className_Full;
            }
        }

        public ControlInfoList? ControlInfos
        {
            get
            {
                return _controlInfos;
            }
        }

        public string? MasterPageTypeName
        {
            get
            {
                return _masterPageTypeName;
            }
        }

        public string? PreviousPageTypeName
        {
            get
            {
                return _previousPageTypeName;
            }
        }

        public UserControlTypeNameDictionary? UserControlTypes
        {
            get
            {
                return _userControls;
            }
        }

        public static int LineOffset
        {
            get
            {
                return 0;
            }
        }

        private Version TargetFrameworkVersion
        {
            get
            {
                if (_targetFrameworkVersion == null)
                {
                    if (_hierarchy != null)
                    {
                        _hierarchy.GetProperty((uint)VSConstants.VSITEMID_ROOT, (int)__VSHPROPID4.VSHPROPID_TargetFrameworkMoniker, out object objTargetFrameworkMoniker);
                        string? targetFrameworkMoniker = objTargetFrameworkMoniker as string;
                        if (!Strings.IsNullOrEmpty(targetFrameworkMoniker))
                        {
                            FrameworkName frameworkName = new FrameworkName(targetFrameworkMoniker);
                            _targetFrameworkVersion = frameworkName.Version;
                        }
                    }
                    if (_targetFrameworkVersion == null)
                    {
                        _targetFrameworkVersion = s_latestFramework;
                    }
                }
                return _targetFrameworkVersion;
            }
        }

        private static Version CurrentFrameworkVersion()
        {
            Version version = typeof(Page).Assembly.GetName().Version;
            return new Version(version.Major, version.Minor);
        }

        public void BeginParse(string? virtualPath, string text)
        {
            if (virtualPath == null)
            {
                throw new ArgumentNullException(nameof(virtualPath));
            }

            // Clean up previous parse state
            EndParse();

            // Make the virtual path absolute
            _virtualPath = VirtualPathUtility.Combine(_appVirtualPath, virtualPath);

            _text = text;

            // Get the contents from the file if not passed in
            if (string.IsNullOrEmpty(_text))
            {
                _text = ReadFileFromVirtualPath(_virtualPath);
            }
        }

        public void EndParse()
        {
            // Clean up BeginParse state
            _virtualPath = null;
            _text = null;

            // Clean up Parse state
            _className_Full = null;
            _className_Name = null;
            _className_Namespace = null;
            _masterPageTypeName = null;
            _previousPageTypeName = null;
            _userControls = null;
            _configRegisterDirectives = null;
            _controlInfos = null;
            _registerDirectives = null;
            _targetFrameworkVersion = null;
        }

        public void Parse()
        {
            _userControls = new UserControlTypeNameDictionary();
            ParseConfig();
            _controlInfos = new ControlInfoList();

            _registerDirectives = new DirectiveList();
            Directive asp = new Directive();
            asp.Attributes["tagprefix"] = "asp";
            asp.Attributes["namespace"] = "System.Web.UI.WebControls";
            asp.Attributes["assembly"] = "System.Web";
            _registerDirectives.Add(asp);
            Directive mobile = new Directive();
            mobile.Attributes["tagprefix"] = "mobile";
            mobile.Attributes["namespace"] = "System.Web.UI.MobileControls";
            mobile.Attributes["assembly"] = "System.Web.Mobile";
            _registerDirectives.Add(mobile);
            if (_configRegisterDirectives != null)
            {
                _registerDirectives.AddRange(_configRegisterDirectives);
            }

            _className_Full = null;
            _className_Name = null;
            _className_Namespace = null;

            ASPTreeParser aspParser = new ASPTreeParser();
            aspParser.BeginParse(TargetFrameworkVersion, _text ?? string.Empty);
            ASPElement? element = aspParser.ParseElement();
            while (element != null)
            {
                if (element.Type == ASPElementType.Directive)
                {
                    if (string.IsNullOrEmpty(element.LCaseName) || element.LCaseName == "page" || element.LCaseName == "control" || element.LCaseName == "master")
                    {
                        // Get inherts value
                        string? inherits = element.GetAttributeValue("inherits");

                        // Strip off any assembly info
                        if (!Strings.IsNullOrEmpty(inherits))
                        {
                            int firstComma = inherits.IndexOf(',');
                            if (firstComma > 0)
                            {
                                // Strip off assembly info
                                _className_Full = inherits.Substring(0, firstComma).Trim();
                            }
                            else
                            {
                                _className_Full = inherits.Trim();
                            }
                        }

                        // Break type name into namespace and name
                        if (!Strings.IsNullOrEmpty(_className_Full))
                        {
                            int lastdot = _className_Full.LastIndexOf('.');
                            if (lastdot >= 0)
                            {
                                _className_Name = _className_Full.Substring(lastdot + 1);
                                _className_Namespace = _className_Full.Substring(0, lastdot);
                            }
                            else
                            {
                                _className_Name = _className_Full;
                                _className_Namespace = null;
                            }
                        }
                    }
                    else if (element.LCaseName == "register")
                    {
                        // Add register directive to collection
                        Directive registerDirective = new Directive(element);
                        _registerDirectives.Add(registerDirective);

                        // We only care about user control registrations, which have a src attribute
                        string? src = element.GetAttributeValue("src");
                        if (!Strings.IsNullOrEmpty(src))
                        {
                            string? tagPrefix = element.GetAttributeValue("tagprefix");
                            if (!Strings.IsNullOrEmpty(tagPrefix))
                            {
                                string? tagName = element.GetAttributeValue("tagname");
                                if (!Strings.IsNullOrEmpty(tagName))
                                {
                                    string fullTagName = tagPrefix + ':' + tagName;
                                    string ucTypeName = "System.Web.UI.UserControl";

                                    try
                                    {
                                        string? typename = GetTypeNameFromVirtualPath(src);
                                        if (typename != null)
                                        {
                                            ucTypeName = typename;
                                        }
                                    }
                                    catch
                                    {
                                        // if there is a failure getting the user control type use System.Web.UI.UserControl
                                    }

                                    if (ucTypeName != null)
                                    {
                                        _userControls[fullTagName] = ucTypeName;
                                    }
                                }
                            }
                        }
                    }
                    else if (element.LCaseName == "mastertype")
                    {
                        AssignTypeNameFromElement(element, ref _masterPageTypeName);
                    }
                    else if (element.LCaseName == "previouspagetype")
                    {
                        AssignTypeNameFromElement(element, ref _previousPageTypeName);
                    }
                }

                element = aspParser.ParseElement();
            }

            // DEBUG code to dump tree to output window
            //aspParser.PrintTree();
            if (aspParser.Root != null)
            {
                ProcessControls(aspParser.Root, null, false, false, 0);
            }
        }

        private void ParseConfig()
        {
            if (_configRegisterDirectives == null)
            {
                _configRegisterDirectives = new DirectiveList();

#pragma warning disable RS0030 // Do not used banned APIs
                IWebApplication? webApp = (IWebApplication)_designerHost.GetService(typeof(IWebApplication));
#pragma warning restore RS0030 // Do not used banned APIs
                if (webApp != null)
                {
                    System.Configuration.Configuration config = webApp.OpenWebConfiguration(true /*readonly*/);
                    if (config != null)
                    {
                        PagesSection section = (PagesSection)config.GetSection("system.web/pages");
                        if (section != null)
                        {
                            // TODO: we should fix up user control register directive src to be app relative
                            foreach (TagPrefixInfo tagPrefix in section.Controls)
                            {
                                Directive directive = new Directive();
                                directive.Name = "register";

                                ElementInformation elemInfo = tagPrefix.ElementInformation;
                                foreach (PropertyInformation propInfo in elemInfo.Properties)
                                {
                                    if (propInfo.Type == typeof(string))
                                    {
                                        if (propInfo.ValueOrigin != PropertyValueOrigin.Default)
                                        {
                                            string name = propInfo.Name;
                                            if (!string.IsNullOrEmpty(name))
                                            {
                                                name = name.ToLowerInvariant();
                                                string value = (string)propInfo.Value;
                                                directive.Attributes[name] = value;
                                            }
                                        }
                                    }
                                }
                                _configRegisterDirectives.Add(directive);

                                // If it is a usercontrol resolve the type
                                if (directive.Attributes.TryGetValue("src", out string src)
                                    && directive.Attributes.TryGetValue("tagprefix", out string prefix)
                                    && directive.Attributes.TryGetValue("tagname", out string tagName))
                                {
                                    string fullTagName = prefix + ':' + tagName;
                                    string ucTypeName = "System.Web.UI.UserControl";

                                    try
                                    {
                                        string? typename = GetTypeNameFromVirtualPath(src);
                                        if (typename != null)
                                        {
                                            ucTypeName = typename;
                                        }
                                    }
                                    catch
                                    {
                                        // if there is a failure getting the user control type use System.Web.UI.UserControl
                                    }

                                    if (ucTypeName != null && _userControls != null)
                                    {
                                        _userControls[fullTagName] = ucTypeName;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void AssignTypeNameFromElement(ASPElement element, ref string? typename)
        {
            string? virtualPath = element.GetAttributeValue("virtualpath");
            if (!Strings.IsNullOrEmpty(virtualPath))
            {
                typename = GetTypeNameFromVirtualPath(virtualPath);
            }
            else
            {
                typename = element.GetAttributeValue("typename");
            }
        }

        private void EnsureVirtualPathInApp(string virtualPath)
        {
            // Make sure it's inside the app
            if (!virtualPath.StartsWith(_appVirtualPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, WebResources.Parser_InvalidPath, virtualPath));
            }
        }

        private string MapPath(string virtualPath)
        {
            virtualPath = VirtualPathUtility.ToAbsolute(virtualPath, _appVirtualPath);

            EnsureVirtualPathInApp(virtualPath);

            string relativePath = virtualPath.Substring(_appVirtualPath.Length);

            // Append the relative path to the physical app path
            string fullPath = Path.Combine(_appPhysicalPath, relativePath).Replace('/', '\\');
            if (!File.Exists(fullPath))
            {
                // Try the metabase path.
#pragma warning disable RS0030 // Do not used banned APIs
                IVsHierarchy hier = (IVsHierarchy)_designerHost.GetService(typeof(IVsHierarchy));
#pragma warning restore RS0030 // Do not used banned APIs
                if (hier != null)
                {
                    WAProject proj = WAProject.GetProjectFromIVsHierarchy(hier);
                    if (proj != null)
                    {
                        string pathFromProject = proj.GetPathForVirtualPath(virtualPath);
                        if (!String.IsNullOrEmpty(pathFromProject))
                            fullPath = pathFromProject;
                    }
                }
            }

            return fullPath;
        }

        private string ReadFileFromVirtualPath(string virtualPath)
        {
            string path = MapPath(virtualPath);

            // REVIEW: do we need to be careful about encoding here?
            using (StreamReader reader = new StreamReader(path))
            {
                return reader.ReadToEnd();
            }
        }

        // Extract a page's type (aspx, ascx or master) name from its inherits attribute
        private string? GetTypeNameFromVirtualPath(string virtualPath)
        {
            virtualPath = VirtualPathUtility.Combine(_virtualPath, virtualPath);
            string text = ReadFileFromVirtualPath(virtualPath);

            string? className = null;

            var directiveParser = new DirectiveParser(text);

            for (; ; )
            {
                // Get the next directive
                Directive? directive = directiveParser.ParseNextDirective();
                if (directive == null)
                    break;

                // We only care about the main directive
                if (!directive.IsMain)
                    continue;

                // Get type name from inherits
                directive.Attributes.TryGetValue("inherits", out string? inherits);
                if (!Strings.IsNullOrEmpty(inherits))
                {
                    int index = inherits.IndexOf(',');
                    if (index > 0)
                    {
                        // Strip off assembly info
                        className = inherits.Substring(0, index).Trim();
                    }
                    else
                    {
                        className = inherits.Trim();
                    }
                }

                // Trim off dot prefix if present
                if (!Strings.IsNullOrEmpty(className) && className.Length > 1 && className.StartsWith(".", StringComparison.Ordinal))
                {
                    className = className.Substring(1);
                }

                // DevDivBugs:29272
                // Special case for predicting class name in VB 
                // when dependent file has not yet been converted.
                if (!Strings.IsNullOrEmpty(className))
                {
                    directive.Attributes.TryGetValue("codefile", out string? codefile);
                    if (!Strings.IsNullOrEmpty(codefile))
                    {
                        directive.Attributes.TryGetValue("codebehind", out string? codebehind);
                        if (Strings.IsNullOrEmpty(codebehind))
                        {
                            string? extension = null;
                            try
                            {
                                extension = Path.GetExtension(codefile);
                            }
                            catch
                            {
                            }

                            if (!Strings.IsNullOrEmpty(extension) && string.Compare(extension, ".vb", StringComparison.OrdinalIgnoreCase) == 0)
                            {
                                IVsHierarchy? hierarchy = CodeGenUtils.GetService<IVsHierarchy>(_designerHost);
                                if (hierarchy != null)
                                {
                                    VsHierarchyItem item = new VsHierarchyItem(hierarchy);
                                    string? defaultNamespace = item.DefaultNamespace;
                                    if (!Strings.IsNullOrEmpty(defaultNamespace))
                                    {
                                        className = defaultNamespace + "." + className;
                                    }
                                }
                            }
                        }
                    }
                }

                return className;
            }

            return null;
        }

        // Note that in general this code should look very similar to that in ndp\fx\src\xsp\system\web\ui\htmltagnametotypemapper.cs, at least in how it
        // maps the control types. 
        private void GetControlTypeInfo(ASPTreeElement element, out string? typeName, out bool isUserControl, out bool isCustomControl, out bool isHtmlControl, out Type? controlType, out Directive? resolvedRegisterDirective)
        {
            typeName = null;
            isUserControl = false;
            isCustomControl = false;
            isHtmlControl = false;
            controlType = null;
            resolvedRegisterDirective = null;

            // Get the tag name
            string? tagName = element.Name?.ToString()?.Trim();

            // Split the tag name into namespace and name
            string? tagName_Namespace = null;
            string? tagName_Name = null;
            if (!Strings.IsNullOrEmpty(tagName))
            {
                int firstColon = tagName.IndexOf(':');
                if (firstColon > 0)
                {
                    tagName_Namespace = tagName.Substring(0, firstColon).Trim();
                    tagName_Name = tagName.Substring(firstColon + 1).Trim();
                }
                else
                {
                    tagName_Name = tagName.Trim();
                }
            }

            if (!Strings.IsNullOrEmpty(tagName_Namespace))
            {
                // Resolve to User Control or Custom Control
                if (_registerDirectives != null)
                {
                    // Iterate through directives to find first match. We want to check against all user controls first as this
                    // avoids calls into the type resolution service to look for types that will never exist in a reference
                    // assembly. This looking for types could cause additional assemblies to be loaded that we don't need.
                    foreach (Directive registerDirective in _registerDirectives)
                    {
                        registerDirective.Attributes.TryGetValue("tagprefix", out string? tagPrefix);

                        if (!Strings.IsNullOrEmpty(tagPrefix) && string.Compare(tagPrefix, tagName_Namespace, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            registerDirective.Attributes.TryGetValue("tagname", out string? tn);
                            registerDirective.Attributes.TryGetValue("src", out string? src);
                            registerDirective.Attributes.TryGetValue("namespace",  out string? ns);
                            registerDirective.Attributes.TryGetValue("assembly", out string? asm);

                            // Check for user control
                            if (!Strings.IsNullOrEmpty(src)
                                && !Strings.IsNullOrEmpty(tn)
                                && Strings.IsNullOrEmpty(ns)
                                && Strings.IsNullOrEmpty(asm)
                                && string.Compare(tn, tagName_Name, StringComparison.OrdinalIgnoreCase) == 0)
                            {
                                isUserControl = true;
                                typeName = GetUserControlTypeName(tagPrefix, tn);
                                resolvedRegisterDirective = registerDirective;
                                return;
                            }
                        }
                    }

                    // Didn't find a user control, so we now have to try the other assemblies.
                    foreach (Directive registerDirective in _registerDirectives)
                    {
                        registerDirective.Attributes.TryGetValue("tagprefix", out string? tagPrefix);

                        if (!Strings.IsNullOrEmpty(tagPrefix) && string.Compare(tagPrefix, tagName_Namespace, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            registerDirective.Attributes.TryGetValue("tagname", out string? tn);
                            registerDirective.Attributes.TryGetValue("src", out string? src);
                            registerDirective.Attributes.TryGetValue("namespace",  out string? ns);
                            registerDirective.Attributes.TryGetValue("assembly", out string? asm);

                            // Skip user controls since we already looked there
                            if (!Strings.IsNullOrEmpty(src)
                                && !Strings.IsNullOrEmpty(tn)
                                && Strings.IsNullOrEmpty(ns)
                                && Strings.IsNullOrEmpty(asm))
                            {
                                continue;
                            }

                            // Check for custom control
                            if (!Strings.IsNullOrEmpty(ns)
                                && Strings.IsNullOrEmpty(src)
                                && Strings.IsNullOrEmpty(tn))
                            {
                                isCustomControl = true;

                                // Save the output type name for the declaration
                                typeName = ns + "." + tagName_Name;

                                // Get the assembly name and strong name
                                string? asm_Name = null;
                                string? asm_StrongName = null;
                                if (!Strings.IsNullOrEmpty(asm))
                                {

                                    int firstComma = asm.IndexOf(',');
                                    if (firstComma > 0)
                                    {
                                        asm_Name = asm.Substring(0, firstComma).Trim();
                                        asm_StrongName = asm.Trim();
                                    }
                                    else
                                    {
                                        asm_Name = asm.Trim();
                                    }
                                }

                                // Try to load the control type
                                if (!Strings.IsNullOrEmpty(asm_Name))
                                {
                                    if (!Strings.IsNullOrEmpty(asm_StrongName))
                                    {
                                        // Try strong assembly qualified type name
                                        controlType = GetType(typeName + ", " + asm_StrongName);
                                    }

                                    if (controlType == null)
                                    {
                                        // Try assembly qualified type name
                                        controlType = GetType(typeName + ", " + asm_Name);
                                    }
                                }
                                else
                                {
                                    // Try non-assembly qualified type name
                                    controlType = GetType(typeName);
                                }

                                if (controlType != null)
                                {
                                    // Get the proper casing of the type name for the declaration
                                    typeName = controlType.FullName;
                                    resolvedRegisterDirective = registerDirective;
                                    return;
                                }

                                // we have a typeName but failed to load the type
                                // continue looping to see if there is another namespace for this tag prefix
                                // if we never load a type the last typeName will be returned
                            }
                        }
                    }
                }
            }
            else if (!Strings.IsNullOrEmpty(tagName_Name))
            {
                // Resolve to Html Control

                isHtmlControl = true;

                if (_htmlTagMap == null)
                {
                    Hashtable t = new Hashtable(10, StringComparer.OrdinalIgnoreCase);
                    t.Add("a", typeof(HtmlAnchor));
                    t.Add("button", typeof(HtmlButton));
                    t.Add("form", typeof(HtmlForm));
                    t.Add("head", typeof(HtmlHead));
                    t.Add("img", typeof(HtmlImage));
                    t.Add("textarea", typeof(HtmlTextArea));
                    t.Add("select", typeof(HtmlSelect));
                    t.Add("table", typeof(HtmlTable));
                    t.Add("tr", typeof(HtmlTableRow));
                    t.Add("td", typeof(HtmlTableCell));
                    t.Add("th", typeof(HtmlTableCell));

                    // 4.5 specific HTML controls
                    if (TargetFrameworkVersion >= s_framework45)
                    {
                        t.Add("audio", typeof(HtmlAudio));
                        t.Add("video", typeof(HtmlVideo));
                        t.Add("track", typeof(HtmlTrack));
                        t.Add("source", typeof(HtmlSource));
                        t.Add("iframe", typeof(HtmlIframe));
                        t.Add("embed", typeof(HtmlEmbed));
                        t.Add("area", typeof(HtmlArea));
                        t.Add("html", typeof(HtmlElement));
                    }
                    _htmlTagMap = t;
                }

                if (_htmlInputTypes == null)
                {
                    Hashtable t = new Hashtable(10, StringComparer.OrdinalIgnoreCase);
                    t.Add("text", typeof(HtmlInputText));
                    t.Add("password", typeof(HtmlInputPassword));
                    t.Add("button", typeof(HtmlInputButton));
                    t.Add("submit", typeof(HtmlInputSubmit));
                    t.Add("reset", typeof(HtmlInputReset));
                    t.Add("image", typeof(HtmlInputImage));
                    t.Add("checkbox", typeof(HtmlInputCheckBox));
                    t.Add("radio", typeof(HtmlInputRadioButton));
                    t.Add("hidden", typeof(HtmlInputHidden));
                    t.Add("file", typeof(HtmlInputFile));
                    _htmlInputTypes = t;
                }

                if (string.Compare("input", tagName, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    string? type = element.GetAttributeValue("type");
                    if (type == null)
                        type = "text";

                    controlType = (Type)_htmlInputTypes[type];
                    if (controlType != null)
                    {
                        typeName = controlType.FullName;
                    }
                    else
                    {
                        if (TargetFrameworkVersion >= s_framework45)
                        {
                            // HtmlInputGenericControl in .net 4.5

                            // Change code to the following when it will compile in VS
                            controlType = typeof(HtmlInputGenericControl);
                            typeName = controlType.FullName;
                        }
                        else
                        {
                            // Unknown input type in .Net 1.0-4.0
                            isHtmlControl = false;
                        }
                    }
                }
                else
                {
                    controlType = (Type)_htmlTagMap[tagName];
                    if (controlType == null)
                    {
                        // Special case for title, link, and meta
                        // - when in a head with runat=server they are strongly typed
                        // - when in a client head they are typed generic
                        if (IsTitleInServerHead(element))
                        {
                            controlType = typeof(HtmlTitle);
                        }
                        else if (IsLinkInServerHead(element))
                        {
                            controlType = typeof(HtmlLink);
                        }
                        else if (IsMetaInServerHead(element))
                        {
                            controlType = typeof(HtmlMeta);
                        }

                        // All other controls are typed generic
                        if (controlType == null)
                        {
                            controlType = typeof(HtmlGenericControl);
                        }
                    }
                    typeName = controlType.FullName;
                }
            }
        }

        private ControlInfo GetControlInfo(ASPTreeElement element)
        {
            ControlInfo controlInfo = new ControlInfo();

            if (element.HasID)
            {
                controlInfo.ID = element.GetAttributeValue("id");
            }

            GetControlTypeInfo(element, out controlInfo.TypeName, out controlInfo.IsUserControl, out controlInfo.IsCustomControl, out controlInfo.IsHtmlControl, out controlInfo.ControlType, out controlInfo.ResolvedRegisterDirective);

            controlInfo.ParseChildrenAsProperties = false;

            return controlInfo;
        }

        private void GetControlDeclareType(ASPTreeElement element, ControlInfo ci)
        {
            if (ci.DeclareTypeName == null)
            {
                // default to control type
                ci.DeclareTypeName = ci.TypeName;
                ci.DeclareType = ci.ControlType;

                // if there is a control builder try to use it to get the declaration type
                if (ci.IsCustomControl && ci.BuilderType != null && ci.ResolvedRegisterDirective != null)
                {
                    try
                    {

                        string parseText = ci.ResolvedRegisterDirective.ToString();
                        parseText += element.Outer?.ToString();
                        DesignTimeParseData parseData = new DesignTimeParseData(_designerHost, parseText);
                        ControlBuilder builder = (ControlBuilder)DesignTimeTemplateParser.ParseTemplate(parseData);
                        ControlBuilder? subBuilder = ControlBuilderWrapper.GetFirstSubBuilder(builder);
                        if (subBuilder != null)
                        {
                            ci.DeclareType = subBuilder.DeclareType;
                            ci.DeclareTypeName = ci.DeclareType.FullName;
                        }
                    }
                    catch (Exception)
                    {
                        // if the builder fails we use the control type
                    }
                }
            }
        }

        private void ProcessControls(ASPTreeElement element, ControlInfo? parentCI, bool ignoreUnknownContent, bool isProperty, int level)
        {
            // Process the element
            ControlInfo? ci = ProcessControl(element, parentCI, ignoreUnknownContent, isProperty, level);

            if (ci != null && ci.ParseChildrenAsProperties)
            {
                // Process the properties
                ProcessControlProperties(element, ci, level);
            }
            else
            {
                // Process the children
                foreach (ASPTreeElement childElement in element.Children)
                {
                    ProcessControls(childElement, parentCI, ignoreUnknownContent, false, level + 1);
                }
            }
        }

        private ControlInfo? ProcessControl(ASPTreeElement element, ControlInfo? parentCI, bool ignoreUnknownContent, bool isProperty, int level)
        {
            ControlInfo? ci = null;

            // DEBUG code to dump process tree
            //ASPTextSpan name = element.Name;
            //if (name != null)
            //{
            //    Debug.WriteLine(string.Empty.PadLeft(level, '*') + name.ToString());
            //}

            if (element.Type == ASPElementType.Tag
                && !isProperty
                && ((parentCI != null && parentCI.ParseChildrenAsProperties) || element.IsRunAtServer || IsAutoServerElementInServerHead(element)))
            {
                ci = GetControlInfo(element);
                bool ignoreContent = ignoreUnknownContent && ci.IsHtmlControl && ci.TypeName == typeof(HtmlGenericControl).FullName;
                if (parentCI != null && ci != null && !ignoreContent)
                {
                    ci.ParseChildrenAsProperties = parentCI.ParseChildrenAsProperties;
                }

                if (!ignoreContent && ci != null &&
                    !Strings.IsNullOrEmpty(ci.TypeName))
                {

                    // Ensure type
                    if (ci.ControlType == null)
                    {
                        ci.ControlType = GetType(ci.TypeName);
                    }

                    // Content control is a special case where
                    // we never want to declare the control and
                    // we always want to process the children
                    // it is runat server and INamingContainer and 
                    // so we need to ignore it.

                    if (ci.ControlType == null || !typeof(System.Web.UI.WebControls.Content).IsAssignableFrom(ci.ControlType))
                    {
                        // Add to output list
                        if (element.HasID && !string.IsNullOrEmpty(ci.ID))
                        {
                            if (_controlInfos == null)
                            {
                                _controlInfos = new ControlInfoList();
                            }

                            _controlInfos.Add(ci);

                            // DEBUG code to dump process tree IDs
                            //Debug.WriteLine(string.Empty.PadLeft(level, '.') + ci.ID);
                        }

                        if (ci.ControlType != null)
                        {
                            System.ComponentModel.AttributeCollection? attColl = null;
                            attColl = TypeDescriptor.GetAttributes(ci.ControlType);
                            if (attColl != null)
                            {
                                // Check ParseChildrenAttribute
                                ParseChildrenAttribute? pca = null;
                                foreach (Attribute attr in attColl)
                                {
                                    if (attr is ParseChildrenAttribute)
                                    {
                                        pca = attr as ParseChildrenAttribute;
                                        if (pca != null && pca.ChildrenAsProperties)
                                        {
                                            ci.ParseChildrenAsProperties = true;
                                            ci.ParseChildrenPropertyName = pca.DefaultProperty;
                                        }
                                        break;
                                    }
                                }

                                // Check ControlBuilderAttribute
                                ControlBuilderAttribute? cba = null;
                                foreach (Attribute attr in attColl)
                                {
                                    if (attr is ControlBuilderAttribute)
                                    {
                                        cba = attr as ControlBuilderAttribute;
                                        if (cba != null)
                                        {
                                            ci.BuilderType = cba.BuilderType;
                                        }
                                        break;
                                    }
                                }
                            }
                        }

                        // Caculate the control declaration type (ci.DeclareTypeName)
                        GetControlDeclareType(element, ci);
                    }
                }
            }

            return ci;
        }

        private void ProcessControlProperties(ASPTreeElement element, ControlInfo ci, int level)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(ci.ControlType);
            if (properties != null)
            {
                if (!string.IsNullOrEmpty(ci.ParseChildrenPropertyName))
                {
                    ProcessControlProperty(element, ci.ParseChildrenPropertyName, true, properties, ci, level + 1);
                }
                else
                {
                    ProcessControlProperties(element, ci, properties, level);
                }
            }
        }

        private void ProcessControlProperties(ASPTreeElement element, ControlInfo ci, PropertyDescriptorCollection properties, int level)
        {
            if (properties != null)
            {
                foreach (ASPTreeElement propertyElement in element.Children)
                {
                    if (propertyElement.Type == ASPElementType.Tag)
                    {
                        ASPTextSpan? name = propertyElement.Name;
                        if (name != null)
                        {
                            string propertyName = name.ToString();
                            if (!Strings.IsNullOrEmpty(propertyName))
                            {
                                ProcessControlProperty(propertyElement, propertyName, false, properties, ci, level + 1);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Examines the childElement to determine:
        ///     - if it represents a control property
        ///     - if the property is a complex type
        ///     - if the property is a collection
        ///     - if the property is a template
        ///     - if the template is a single instance template
        /// 
        /// Then processes the element appropriately
        /// </summary>
        private void ProcessControlProperty(ASPTreeElement childElement, string? propertyName, bool isDefaultProperty, PropertyDescriptorCollection properties, ControlInfo ci, int level)
        {
            bool isTemplate = false;
            bool isSingleInstanceTemplate = false;
            bool isProperty = false;
            bool isComplex = false;
            bool isCollection = false;
            bool ignoreUnknownContent = false;
            PropertyDescriptor? prop = null;

            if (childElement.Type == ASPElementType.Tag)
            {
                if (!Strings.IsNullOrEmpty(propertyName))
                {
                    // DEBUG code to dump process tree
                    //Debug.WriteLine(string.Empty.PadLeft(level, '+') + propertyName);

                    prop = properties.Find(propertyName, true);
                    if (prop != null)
                    {
                        isProperty = true;

                        if (typeof(object).IsAssignableFrom(prop.PropertyType) && prop.PropertyType != typeof(string))
                        {
                            isComplex = true;
                        }

                        if (typeof(ITemplate).IsAssignableFrom(prop.PropertyType))
                        {
                            isTemplate = true;

                            System.ComponentModel.AttributeCollection? attColl = prop.Attributes;
                            if (attColl != null)
                            {
                                TemplateInstanceAttribute? tia = (TemplateInstanceAttribute)attColl[typeof(TemplateInstanceAttribute)];
                                if (tia != null)
                                {
                                    if (tia.Instances == TemplateInstance.Single)
                                    {
                                        isSingleInstanceTemplate = true;
                                    }
                                }
                            }
                        }
                        else if (typeof(ICollection).IsAssignableFrom(prop.PropertyType))
                        {
                            isCollection = true;

                            System.ComponentModel.AttributeCollection? attColl = null;
                            attColl = prop.Attributes;
                            if (attColl != null)
                            {
                                foreach (Attribute attr in attColl)
                                {
                                    if (attr.GetType().FullName == "System.Web.UI.IgnoreUnknownContentAttribute")
                                    {
                                        ignoreUnknownContent = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // If the property is a multi-instance template we DO NOT want to 
            // process the controls in the template.
            // If the property is a single-instance template we DO want to 
            // process the controls in the template.
            // If the property is not a template we process it.
            // If the element is not a property we process it.
            // If the element is a collection we process it
            // If the element is a complex property we process it
            // If the property is not complex we do not process it
            if (isProperty && isComplex && !isCollection && !isTemplate)
            {
                if (prop != null)
                {
                    ProcessControlProperties(childElement, ci, TypeDescriptor.GetProperties(prop.PropertyType), level);
                }
            }
            else if (isProperty && isCollection)
            {
                ProcessControls(childElement, ci, ignoreUnknownContent, true, level);
            }
            else if (isProperty && isSingleInstanceTemplate)
            {
                ProcessControls(childElement, null, false, true, level);
            }
            else if (!isProperty)
            {
                ProcessControls(childElement, null, false, isDefaultProperty, level);
            }
        }

        public string GetUserControlTypeName(string tagPrefix, string tagName)
        {
            string? typeName = null;
            _userControls?.TryGetValue(tagPrefix + ":" + tagName, out typeName);
            if (Strings.IsNullOrEmpty(typeName))
            {
                typeName = "System.Web.UI.UserControl";
            }
            return typeName;
        }

        protected Type? GetType(string typeName)
        {
            try
            {
                return _designerHost.GetType(typeName);
            }
            catch
            {
            }

            return null;
        }

        private static bool IsAutoServerElementInServerHead(ASPTreeElement element)
        {
            if (IsTitleInServerHead(element)
                || IsLinkInServerHead(element)
                || IsMetaInServerHead(element))
            {
                return true;
            }

            return false;
        }

        private static bool IsTitleInServerHead(ASPTreeElement element)
        {
            if (element.LCaseName == "title" && IsInHeadRunAtServer(element))
            {
                return true;
            }
            return false;
        }

        private static bool IsLinkInServerHead(ASPTreeElement element)
        {
            if (element.LCaseName == "link" && IsInHeadRunAtServer(element))
            {
                return true;
            }
            return false;
        }

        private static bool IsMetaInServerHead(ASPTreeElement element)
        {
            if (element.LCaseName == "meta" && IsInHeadRunAtServer(element))
            {
                return true;
            }
            return false;
        }
        private static bool IsInHeadRunAtServer(ASPTreeElement? element)
        {
            while (element != null)
            {
                element = element.Parent;
                if (element != null && element.IsRunAtServer && element.LCaseName == "head")
                {
                    return true;
                }
            }

            return false;
        }
    }

    internal class UserControlTypeNameDictionary : Dictionary<string, string>
    {
        public UserControlTypeNameDictionary() : base(StringComparer.OrdinalIgnoreCase)
        {
        }
    }

    internal class ControlInfo
    {
        public string? ID;
        public string? TypeName;
        public Type? ControlType;
        public Type? BuilderType;
        public Type? DeclareType;
        public string? DeclareTypeName;
        public bool IsUserControl;
        public bool IsCustomControl;
        public bool IsHtmlControl;
        public bool ParseChildrenAsProperties;
        public string? ParseChildrenPropertyName;
        public Directive? ResolvedRegisterDirective;
    }

    internal class ControlInfoList : List<ControlInfo>
    {
    }
}
