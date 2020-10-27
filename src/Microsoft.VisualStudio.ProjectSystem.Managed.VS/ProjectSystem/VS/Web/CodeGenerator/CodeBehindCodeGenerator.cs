// Licensed to the .NET Foundation under one or more agreements. The .NET Foundation licenses this file to you under the MIT license. See the LICENSE.md file in the project root for more information.

using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Formatting;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Designer.Interfaces;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Design.Serialization;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Web.Application;
using CodeDomNamespace = System.CodeDom.CodeNamespace;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    internal class CodeBehindCodeGenerator : IVsCodeBehindCodeGenerator
    {

        private static readonly Guid GUID_CSHARP_LANG_SERVICE = new Guid("694DD9B6-B865-4C5B-AD85-86356E9C88DC");
        private static readonly Guid GUID_VB_LANG_SERVICE = new Guid("E34ACDC0-BAAE-11D0-88BF-00A0C9110049");

        private ServiceProvider? _serviceProvider;
        private IVsHierarchy? _hierarchy;
        private CodeGeneratorOptions? _codeGeneratorOptions;

        // Generate state
        private VsHierarchyItem? _itemCode;
        private VsHierarchyItem? _itemDesigner;
        private CodeDomProvider? _codeDomProvider;
        private CodeCompileUnit? _codeCompileUnit;
        private CodeDomNamespace? _codeNamespace;
        private CodeTypeDeclaration? _codeTypeDeclaration;
        private bool _create;
        private FieldDataDictionary? _codeFields;
        private FieldDataDictionary? _designerFields;
        private string? _className_Full;
        private string? _className_Namespace;
        private string? _className_Name;

        ///-------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Constructor
        /// </summary>
        ///-------------------------------------------------------------------------------------------------------------
        public CodeBehindCodeGenerator()
        {
        }

        ///-------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Finalizer
        /// </summary>
        ///-------------------------------------------------------------------------------------------------------------
        ~CodeBehindCodeGenerator()
        {
            System.Diagnostics.Debug.Fail("CodeBehindCodeGenerator was not disposed.");
            Dispose();
        }

        ///-------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Initializes the generator state.
        /// </summary>
        ///-------------------------------------------------------------------------------------------------------------
        void IVsCodeBehindCodeGenerator.Initialize(IVsHierarchy hierarchy)
        {
            _hierarchy = hierarchy;
            _hierarchy.GetSite(out IOleServiceProvider serviceProvider);
            _serviceProvider = new ServiceProvider(serviceProvider);
            _codeGeneratorOptions = new CodeGeneratorOptions();
        }

        private IVsEditorAdaptersFactoryService? _editorAdapterService;
        private IVsEditorAdaptersFactoryService? EditorAdapterService
        {
            get
            {
                if (_editorAdapterService == null && _serviceProvider != null)
                {

                    var componentModel = _serviceProvider.GetService(typeof(SComponentModel)) as IComponentModel;
                    _editorAdapterService = componentModel?.DefaultExportProvider?.GetExport<IVsEditorAdaptersFactoryService>()?.Value;
                }

                return _editorAdapterService;
            }
        }

        void IVsCodeBehindCodeGenerator.Close()
        {
            Dispose();
        }

        public virtual void Dispose()
        {
            if (_serviceProvider != null)
            {
                _serviceProvider.Dispose();
                _serviceProvider = null;
            }

            _hierarchy = null;
            _codeGeneratorOptions = null;

            DisposeGenerateState();

            GC.SuppressFinalize(this);
        }

        protected virtual void DisposeGenerateState()
        {
            try
            {
                _itemCode = null;
                _itemDesigner = null;
                _codeCompileUnit = null;
                _codeNamespace = null;
                _codeTypeDeclaration = null;
                _create = false;
                _codeFields = null;
                _designerFields = null;
                _className_Full = null;
                _className_Namespace = null;
                _className_Name = null;
                if (_codeDomProvider != null)
                {
                    _codeDomProvider.Dispose();
                    _codeDomProvider = null;
                }
            }
            catch
            {
            }
        }

        protected string? ClassName_Full
        {
            get
            {
                return _className_Full;
            }
        }

        protected string? ClassName_Namespace
        {
            get
            {
                return _className_Namespace;
            }
        }

        protected string? ClassName_Name
        {
            get
            {
                return _className_Name;
            }
        }

        protected VsHierarchyItem? ItemCode
        {
            get
            {
                return _itemCode;
            }
        }

        protected CodeDomProvider? CodeDomProvider
        {
            get
            {
                return _codeDomProvider;
            }
        }

        protected CodeGeneratorOptions? CodeGeneratorOptions
        {
            get
            {
                return _codeGeneratorOptions;
            }
        }

        protected CodeCompileUnit? CodeCompileUnit
        {
            get
            {
                return _codeCompileUnit;
            }
        }

        protected CodeDomNamespace? CodeNamespace
        {
            get
            {
                return _codeNamespace;
            }
        }

        protected CodeTypeDeclaration? CodeTypeDeclaration
        {
            get
            {
                return _codeTypeDeclaration;
            }
        }

        /// <summary>
        /// Create CodeDomProvider for the language of the file
        /// </summary>
        protected virtual CodeDomProvider? CreateCodeDomProvider(uint itemid)
        {
            if (_serviceProvider?.GetService(typeof(IVSMDDesignerService)) is IVSMDCodeDomCreator vsmdCodeDomCreator)
            {
                IVSMDCodeDomProvider vsmdCodeDomProvider = vsmdCodeDomCreator.CreateCodeDomProvider(_hierarchy, (int)itemid);
                if (vsmdCodeDomProvider != null)
                {
                    if (vsmdCodeDomProvider.CodeDomProvider is CodeDomProvider codeDomProvider)
                    {
                        return codeDomProvider;
                    }
                }
            }

            System.Diagnostics.Debug.Fail("Failed to create CodeDomProvider");
            return null;
        }

        /// <summary>
        /// Returns field names in the specified class using code model.
        /// If publicOnly is true only public fields are returned.
        /// </summary>
        protected FieldDataDictionary? GetFieldNames(VsHierarchyItem itemCode, string className, bool caseSensitive, bool onlyBaseClasses, int maxDepth)
        {
            FieldDataDictionary? fields = null;

            if (itemCode != null)
            {
                CodeClass? codeClass = FindClass(itemCode, className);
                if (codeClass != null)
                {
                    GetFieldNames(codeClass, caseSensitive, onlyBaseClasses, 0, maxDepth, ref fields);
                }
            }

            return fields;
        }

        /// <summary>
        /// Returns field names in the specified class using code model.
        /// If publicOnly is true only public fields are returned.
        /// </summary>
        protected void GetFieldNames(CodeClass codeClass, bool caseSensitive, bool onlyBaseClasses, int depth, int maxDepth, ref FieldDataDictionary? fields)
        {
            if (codeClass != null && depth <= maxDepth)
            {
                if (!(onlyBaseClasses && depth == 0))
                {
                    foreach (CodeElement codeElement in codeClass.Members)
                    {
                        //vsCMElement kind = codeElement.Kind; //vsCMElementVariable
                        if (codeElement is CodeVariable codeVariable)
                        {
                            var field = new FieldData(codeClass, codeVariable, depth);

                            if (!string.IsNullOrEmpty(field.Name))
                            {
                                if (fields == null)
                                {
                                    fields = new FieldDataDictionary(caseSensitive);
                                }

                                try
                                {
                                    if (!fields.ContainsKey(field.Name))
                                    {
                                        fields.Add(field.Name, field);
                                    }
                                }
                                catch
                                {
                                }
                            }
                        }
                    }
                }

                foreach (CodeElement baseCodeElement in codeClass.Bases)
                {
                    // Make sure the base class isn't us. If so we ignore it (see DevDiv #481698)
                    if (baseCodeElement is CodeClass baseCodeClass && string.Compare(baseCodeClass.FullName, codeClass.FullName, StringComparison.Ordinal) != 0)
                    {
                        CodeElements? partCodeElements = null;
                        if (baseCodeClass is CodeClass2 baseCodeClass2)
                        {
                            vsCMClassKind classKind = baseCodeClass2.ClassKind;
                            if ((classKind | vsCMClassKind.vsCMClassKindPartialClass) == vsCMClassKind.vsCMClassKindPartialClass)
                            {
                                try
                                {
                                    partCodeElements = baseCodeClass2.Parts;
                                }
                                catch
                                {
                                }
                            }
                        }

                        if (partCodeElements != null && partCodeElements.Count > 1)
                        {
                            foreach (CodeElement partCodeElement in partCodeElements)
                            {
                                if (partCodeElement is CodeClass partCodeClass)
                                {
                                    GetFieldNames(partCodeClass, caseSensitive, onlyBaseClasses, depth + 1, maxDepth, ref fields);
                                }
                            }
                        }
                        else
                        {
                            GetFieldNames(baseCodeClass, caseSensitive, onlyBaseClasses, depth + 1, maxDepth, ref fields);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Locates the code model CodeClass 
        /// </summary>
        protected virtual CodeClass? FindClass(VsHierarchyItem item, string className)
        {
            if (item != null)
            {
                try
                {
                    ProjectItem? projectItem = item.ProjectItem;
                    if (projectItem != null)
                    {
                        FileCodeModel fileCodeModel = projectItem.FileCodeModel;
                        if (fileCodeModel != null)
                        {
                            return FindClass(fileCodeModel.CodeElements, className);
                        }
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        /// <summary>
        /// Recursively searches the CodeElements for the specified class
        /// </summary>
        protected virtual CodeClass? FindClass(CodeElements codeElements, string className)
        {
            if (codeElements != null && !string.IsNullOrEmpty(className))
            {
                foreach (CodeElement codeElement in codeElements)
                {
                    vsCMElement kind = codeElement.Kind;
                    if (kind == vsCMElement.vsCMElementClass)
                    {
                        if (codeElement is CodeClass codeClass && string.Compare(codeClass.FullName, className, StringComparison.Ordinal) == 0)
                        {
                            return codeClass;
                        }
                    }
                    else if (kind == vsCMElement.vsCMElementNamespace)
                    {
                        if (codeElement is EnvDTE.CodeNamespace codeNamespace)
                        {
                            CodeClass? codeClass = FindClass(codeNamespace.Children, className);
                            if (codeClass != null)
                            {
                                return codeClass;
                            }
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Gets a VshierarchyItem for the designer file if possible.
        /// Will create new file if specified and not existing.
        /// </summary>
        protected virtual VsHierarchyItem? GetDesignerItem(VsHierarchyItem itemCode, bool create)
        {
            VsHierarchyItem? itemDesigner = null;

            if (itemCode != null && _hierarchy != null)
            {
                // Calculate codebehind and designer file paths 
                string codeBehindFile = itemCode.FullPath;
                string? designerFile = null;
                if (!string.IsNullOrEmpty(codeBehindFile))
                {
                    designerFile = codeBehindFile.Insert(codeBehindFile.LastIndexOf(".", StringComparison.Ordinal), ".designer");
                }

                // Try to locate existing designer file
                if (!Strings.IsNullOrEmpty(designerFile))
                {
                    itemDesigner = VsHierarchyItem.CreateFromMoniker(designerFile, _hierarchy);
                    if (itemDesigner != null)
                    {
                        return itemDesigner;
                    }
                }

                // Create empty designer file if requested
                if (create && !Strings.IsNullOrEmpty(designerFile))
                {
                    ProjectItem? projectItemCode = itemCode.ProjectItem;
                    if (projectItemCode != null)
                    {
                        ProjectItems? projectItems = projectItemCode.Collection;
                        if (projectItems != null)
                        {
                            try
                            {
                                // This will create a file with UTF8 encoding and BOM.
                                // If you want UTF8 without BOM, specify Encoding.Default instead of Encoding.UTF8
                                using (StreamWriter sw = new StreamWriter(designerFile, false, Encoding.UTF8))
                                {
                                    sw.WriteLine(" ");
                                }

                                projectItems.AddFromFileCopy(designerFile);
                            }
                            catch
                            {
                            }

                            itemDesigner = VsHierarchyItem.CreateFromMoniker(designerFile, _hierarchy);
                            if (itemDesigner != null)
                            {
                                return itemDesigner;
                            }
                            else
                            {
                                try
                                {
                                    File.Delete(designerFile);
                                }
                                catch
                                { }
                            }
                        }
                    }
                }
            }

            return itemDesigner;
        }

        protected static bool IsCaseSensitive(CodeDomProvider codeDomProvider)
        {
            return !((codeDomProvider.LanguageOptions & LanguageOptions.CaseInsensitive) == LanguageOptions.CaseInsensitive);
        }

        bool IVsCodeBehindCodeGenerator.IsGenerateAllowed(string document, string codeBehindFile, bool create)
        {
            if (_hierarchy != null)
            {
                VsHierarchyItem? itemCode = VsHierarchyItem.CreateFromMoniker(codeBehindFile, _hierarchy);
                if (itemCode != null)
                {
                    VsHierarchyItem? itemDesigner = GetDesignerItem(itemCode, false);

                    if ((itemDesigner != null && itemDesigner.IsBuildActionCompile)
                        || (itemDesigner == null && create))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        void IVsCodeBehindCodeGenerator.BeginGenerate(string document, string codeBehindFile, string className_Full, bool create)
        {
            DisposeGenerateState();

            if (_hierarchy == null)
            {
                return;
            }

            _itemCode = VsHierarchyItem.CreateFromMoniker(codeBehindFile, _hierarchy);
            if (_itemCode == null)
            {
                return;
            }

            _itemDesigner = GetDesignerItem(_itemCode, false);
            _create = create;
            _className_Full = className_Full;

            // Break full name into namespace and name
            if (!string.IsNullOrEmpty(_className_Full))
            {
                int lastdot = _className_Full.LastIndexOf('.');
                if (lastdot >= 0)
                {
                    _className_Namespace = _className_Full.Substring(0, lastdot);
                    _className_Name = _className_Full.Substring(lastdot + 1);
                }
                else
                {
                    _className_Namespace = null;
                    _className_Name = _className_Full;
                }
            }

            if (_itemCode != null)
            {
                // Get the CodeDomProvider for the language (MergedCodeDomProvider C#/VB)
                _codeDomProvider = CreateCodeDomProvider(_itemCode.VsItemID);

                if (_codeDomProvider != null && _itemDesigner != null)
                {
                    // Get the field names so we can preserve location and access
                    bool caseSensitive = IsCaseSensitive(_codeDomProvider);

                    _codeFields = GetFieldNames(_itemCode, _className_Full, caseSensitive, false, 30);
                    _designerFields = GetFieldNames(_itemDesigner, _className_Full, caseSensitive, false, 0);

                    // Generate the code objects
                    _codeCompileUnit = GenerateCodeCompileUnit();
                    _codeNamespace = GenerateCodeNamespace();
                    _codeTypeDeclaration = GenerateCodeTypeDeclaration();
                    _codeCompileUnit.Namespaces.Add(_codeNamespace);
                    _codeNamespace.Types.Add(_codeTypeDeclaration);
                }
            }
        }

        protected virtual CodeCompileUnit GenerateCodeCompileUnit()
        {
            return new CodeCompileUnit();
        }

        protected virtual CodeDomNamespace GenerateCodeNamespace()
        {
            // Create namespace and add to compile unit
            string? ns = GetClassNamespace();
            return Strings.IsNullOrEmpty(ns) ? new CodeDomNamespace() : new CodeDomNamespace(ns);
        }

        protected virtual CodeTypeDeclaration GenerateCodeTypeDeclaration()
        {
            // Create partial class definition
            var ctd = new CodeTypeDeclaration(ClassName_Name);
            ctd.IsPartial = true;

            return ctd;
        }

        void IVsCodeBehindCodeGenerator.EnsureStronglyTypedProperty(string propertyName, string propertyTypeName)
        {
            CodeMemberProperty prop = new CodeMemberProperty();
            prop.Attributes &= ~MemberAttributes.AccessMask;
            prop.Attributes &= ~MemberAttributes.ScopeMask;
            prop.Attributes |= MemberAttributes.Final | MemberAttributes.New | MemberAttributes.Public;
            prop.Name = propertyName;
            prop.Type = new CodeTypeReference(propertyTypeName);

            // Add doc comment
            prop.Comments.Add(new CodeCommentStatement("<summary>", true));
            prop.Comments.Add(new CodeCommentStatement(string.Format(WebResources.Generator_DocCommentSummaryProperty, propertyName), true));
            prop.Comments.Add(new CodeCommentStatement("</summary>", true));
            prop.Comments.Add(new CodeCommentStatement("<remarks>", true));
            prop.Comments.Add(new CodeCommentStatement(WebResources.Generator_DocCommentRemarksProperty, true));
            prop.Comments.Add(new CodeCommentStatement("</remarks>", true));

            CodePropertyReferenceExpression propRef = new CodePropertyReferenceExpression(
                new CodeBaseReferenceExpression(), propertyName);

            prop.GetStatements.Add(new CodeMethodReturnStatement(new CodeCastExpression(propertyTypeName, propRef)));
            CodeTypeDeclaration?.Members?.Add(prop);
        }

        void IVsCodeBehindCodeGenerator.EnsureControlDeclaration(string name, string typeName)
        {
            if (ShouldDeclareField(name, typeName))
            {
                var typeReference = new CodeTypeReference(typeName, CodeTypeReferenceOptions.GlobalReference);
                var field = new CodeMemberField(typeReference, name);

                // Add doc comment
                field.Comments.Add(new CodeCommentStatement("<summary>", true));
                field.Comments.Add(new CodeCommentStatement(string.Format(WebResources.Generator_DocCommentSummaryField, name), true));
                field.Comments.Add(new CodeCommentStatement("</summary>", true));
                field.Comments.Add(new CodeCommentStatement("<remarks>", true));
                field.Comments.Add(new CodeCommentStatement(WebResources.Generator_DocCommentRemarksField1, true));
                field.Comments.Add(new CodeCommentStatement(WebResources.Generator_DocCommentRemarksField2, true));
                field.Comments.Add(new CodeCommentStatement("</remarks>", true));

                // Check if someone made the declaration public in the designer file
                bool isPublic = false;
                if (_designerFields != null && _designerFields.ContainsKey(name))
                {
                    FieldData fieldData = _designerFields[name];
                    if (fieldData != null)
                    {
                        isPublic = fieldData.IsPublic;
                    }
                }

                // Set access to public or protected
                field.Attributes &= ~MemberAttributes.AccessMask;
                if (isPublic)
                {
                    field.Attributes |= MemberAttributes.Public;
                }
                else
                {
                    field.Attributes |= MemberAttributes.Family;
                }

                SetAdditionalFieldData(field);

                CodeTypeDeclaration?.Members?.Add(field);
            }
        }

        void IVsCodeBehindCodeGenerator.Generate()
        {
            DocData? ddDesigner = null;
            DocDataTextWriter? designerWriter = null;

            try
            {
                if (_itemCode != null && _codeDomProvider != null && _serviceProvider != null)
                {
                    // Generate the code
                    string generatedCode = GenerateCode();

                    // Create designer file if requested
                    if (_itemDesigner == null && _create)
                    {
                        _itemDesigner = GetDesignerItem(_itemCode, true);
                    }

                    if (_itemDesigner == null)
                    {
                        return;
                    }

                    // See if generated code changed
                    string designerContents = _itemDesigner.GetDocumentText();
                    if (!BufferEquals(designerContents, generatedCode))  // Would be nice to just compare lengths but the buffer gets formatted after insertion
                    {
                        ddDesigner = new LockedDocData(_serviceProvider, _itemDesigner.FullPath);

                        // Try to check out designer file (this throws)
                        ddDesigner.CheckoutFile(_serviceProvider);

                        // First write the non-formatted text. This is necessary since the formatter appplies changes to the contents
                        // of the text buffer.
                        designerWriter = new DocDataTextWriter(ddDesigner, disposeDocData: false);
                        designerWriter.Write(generatedCode);
                        designerWriter.Flush();
                        designerWriter.Close();

                        if (FormatDocument(ddDesigner.Buffer, generatedCode, out string? formattedText))
                        {
                            // Format document made changes, so write out the formatted code 
                            designerWriter = new DocDataTextWriter(ddDesigner);
                            designerWriter.Write(formattedText);
                            designerWriter.Flush();
                            designerWriter.Close();
                        }
                    }
                }
            }
            finally
            {
                if (designerWriter != null)
                {
                    designerWriter.Dispose();
                }
                if (ddDesigner != null)
                {
                    ddDesigner.Dispose();
                }
            }
        }

        /// <summary>
        /// Formats the document if vb or c#. Returns true if changes were made and formatted contents are returned
        /// in formattedCode
        /// </summary>
        private bool FormatDocument(IVsTextBuffer vsTextBuffer, string generatedCode, [NotNullWhen(returnValue: true)] out string? formattedCode)
        {
            formattedCode = null;
            try
            {
                SyntaxTree? tree = GetSyntaxTree(generatedCode);
                if (tree != null)
                {
                    ITextBuffer? textBuffer = EditorAdapterService?.GetDocumentBuffer(vsTextBuffer);
                    if (textBuffer != null)
                    {
                        ITextSnapshot snapshot = textBuffer.CurrentSnapshot;
                        var workspace = Workspace.GetWorkspaceRegistration(textBuffer.AsTextContainer())?.Workspace;
                        if (workspace != null)
                        {
                            var changes = Formatter.GetFormattedTextChanges(tree.GetRoot(CancellationToken.None), workspace, options: null, CancellationToken.None);
                            if (changes.Count > 0)
                            {
                                formattedCode = ApplyChangesToSnapshot(changes, snapshot);
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Fail($"Formatting threw an exception {ex.Message}");
            }

            return false;
        }

        protected virtual SyntaxTree? GetSyntaxTree(string generatedCode)
        {
            // Default returns null
            return null;
        }

        private static string ApplyChangesToSnapshot(IEnumerable<TextChange> changes, ITextSnapshot snapshot)
        {
            StringBuilder sb = new StringBuilder();
            int start = 0;
            foreach (TextChange textChange in changes)
            {
                sb.Append(snapshot.GetText(start, textChange.Span.Start - start));
                sb.Append(textChange.NewText);
                start = textChange.Span.End;
            }

            sb.Append(snapshot.GetText(start, snapshot.Length - start));

            return sb.ToString();
        }

        /// <summary>
        /// Virtual method to allow language specific generation override
        /// </summary>
        protected virtual string GenerateCode()
        {
            StringWriter stringWriter = new StringWriter(CultureInfo.InvariantCulture);
            _codeDomProvider?.GenerateCodeFromCompileUnit(CodeCompileUnit, stringWriter, CodeGeneratorOptions);
            return stringWriter.ToString();
        }

        /// <summary>
        /// Virtual method to allow language specific determination
        /// </summary>
        protected virtual bool ShouldDeclareField(string name, string typeName)
        {
            // Don't add field if already defined in codebehind or exposed from base class
            bool declareField = true;
            if (_codeFields != null && _codeFields.ContainsKey(name))
            {
                FieldData fieldData = _codeFields[name];
                if (fieldData != null)
                {
                    if (fieldData.Depth == 0)
                    {
                        // For immediate class we don't re-declare regardless
                        // of access modifiers.  If the field is not visible to the run-time
                        // it will not be set and code against it will fail.
                        declareField = false;
                    }
                    else if (fieldData.IsProtected || fieldData.IsPublic)
                    {
                        // For bases we do not re-declare if already accessible
                        // (internal, private are not accesible to page assembly)
                        declareField = false;
                    }
                }
            }

            // If case sensitive and there already is a field in the base class
            // of a different case it will cause an "Ambiguous match" error in ASP.Net
            // when parsing the page
            if (_codeFields != null && !Strings.IsNullOrEmpty(name) && _codeDomProvider != null && IsCaseSensitive(_codeDomProvider))
            {
                foreach (FieldData fieldData in _codeFields.Values)
                {
                    if (fieldData.Depth == 0)
                    {
                        string fieldName = fieldData.Name;
                        if (!string.IsNullOrEmpty(fieldName)
                            && fieldName.Length == name.Length
                            && (string.Compare(fieldName, name, StringComparison.OrdinalIgnoreCase) == 0)
                            && !(string.Compare(fieldName, name, StringComparison.Ordinal) == 0))
                        {
                            throw new Exception(string.Format(WebResources.Generator_CaseConflict, name, fieldName));
                        }
                    }
                }
            }

            return declareField;
        }

        /// <summary>
        /// Virtual method to allow language specific adjustment of root classname namespace
        /// </summary>
        protected virtual string? GetClassNamespace()
        {
            return _className_Namespace;
        }

        /// <summary>
        /// Virtual method to allow language specific adjustment of field
        /// </summary>
        protected virtual void SetAdditionalFieldData(CodeMemberField field)
        {
            return;
        }

        /// <summary>
        /// Compare buffers ignoring whitespace
        /// </summary>
        protected static bool BufferEquals(string str1, string str2)
        {
            int i1 = 0;
            int i2 = 0;
            int len1 = str1.Length;
            int len2 = str2.Length;

            for (; ; )
            {
                while (i1 < len1 && char.IsWhiteSpace(str1[i1]))
                    i1++;

                while (i2 < len2 && char.IsWhiteSpace(str2[i2]))
                    i2++;

                if (i1 >= len1)
                {
                    if (i2 >= len2)
                    {
                        return true; // ended with whitespace
                    }
                    return false;    // str1 ended before str2
                }
                else if (i2 >= len2)
                {
                    return false;    // str2 ended before str1
                }
                else if (str1[i1] != str2[i2])
                {
                    return false;    // different chars
                }

                i1++; i2++;          // advance
            }
        }
    }

    [Guid("349c5859-65df-11da-9384-00065b846f21")]
    internal class WACodeBehindCodeGenerator : CodeBehindCodeGenerator
    {
    }
}
