using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Web;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Web.Application;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    /// <summary>
    ///
    ///     Generator - Generates the .designer file for the specified document when called. 
    ///
    /// </summary>
    [Guid("ecd40787-8671-4666-aeb9-df653c8ab5bc")]
    internal class Generator : IVsCodeBehindCodeGeneratorService
    {
        private ServiceProvider? _serviceProvider;
        private IVsHierarchy? _hierarchy;
        private ParserDesignerHost? _designerHost;
        private Parser? _parser;
        private string? _appVirtualPath;
        private string? _appPhysicalPath;
        private string? _codeExtension;
        private string? _lastDocument;
        private string? _lastDocumentContents;
        private ErrorListProvider? _taskProvider;
        private IVsCodeBehindCodeGenerator? _codeBehindCodeGenerator;
        private bool _fInProcessUpdateDesignerClass;

        /// <summary>
        ///     Constructor
        /// </summary>
        public Generator()
        {
            _fInProcessUpdateDesignerClass = false;
        }

        /// <summary>
        ///     Finalizer
        /// </summary>
        ~Generator()
        {
            System.Diagnostics.Debug.Fail("Generator was not disposed.");
            Dispose();
        }

        int IVsCodeBehindCodeGeneratorService.InitGenerator(IVsHierarchy hierarchy, string appVirtualPath, string appPhysicalPath)
        {
            int hr = HResult.OK;
            try
            {
                Init(hierarchy, appVirtualPath, appPhysicalPath);
            }
            catch
            {
                hr = HResult.Fail;
            }
            return hr;
        }

        int IVsCodeBehindCodeGeneratorService.CloseGenerator()
        {
            int hr = HResult.OK;
            try
            {
                Dispose();
            }
            catch
            {
                hr = HResult.Fail;
            }
            return hr;
        }

        int IVsCodeBehindCodeGeneratorService.UpdateDesignerClass(string document, string codeBehind, string codeBehindFile, bool force, bool create)
        {
            UpdateDesignerClass(document, codeBehind, codeBehindFile, force, create);
            return HResult.OK;
        }

        /// <summary>
        ///     Initializes the generator state.
        /// </summary>
        public virtual void Init(IVsHierarchy hierarchy, string appVirtualPath, string appPhysicalPath)
        {
            // TODO: timmcb move appVirtuapPath appPhysicalPath somewhere and get from hierarchy
            _hierarchy = hierarchy;
            _hierarchy.GetSite(out IOleServiceProvider serviceProvider);
            _serviceProvider = new ServiceProvider(serviceProvider);
            _designerHost = new ParserDesignerHost(_serviceProvider, _hierarchy);
            _appVirtualPath = appVirtualPath;
            _appPhysicalPath = appPhysicalPath;
            _codeExtension = CodeGenUtils.GetWAPLanguageProperty<string>(_hierarchy, "CodeFileExtension", null);
            _parser = new Parser(_appVirtualPath, _appPhysicalPath, _designerHost, _hierarchy);
            _designerHost.SetParser(_parser);
            _taskProvider = new ErrorListProvider(_serviceProvider);
            _codeBehindCodeGenerator = CreateCodeBehindCodeGenerator();
            _fInProcessUpdateDesignerClass = false;
        }

        /// <summary>
        ///     Helper to create registered code behind code generator
        /// </summary>
        private IVsCodeBehindCodeGenerator? CreateCodeBehindCodeGenerator()
        {
            if (_hierarchy != null)
            {
                string? clsid = CodeGenUtils.GetWAPLanguageProperty<string>(_hierarchy, "CodeBehindCodeGenerator", null);
                if (clsid != null && _serviceProvider != null)
                {
                    IVsCodeBehindCodeGenerator? codeBehindCodeGenerator = CodeGenUtils.CreateInstance<IVsCodeBehindCodeGenerator>(clsid);
                    if (codeBehindCodeGenerator != null)
                    {
                        codeBehindCodeGenerator.Initialize(_hierarchy);
                        return codeBehindCodeGenerator;
                    }
                }
            }

            return null;
        }

        /// <summary>
        ///     Cleans up member state.
        /// </summary>
        public virtual void Dispose()
        {
            if (_taskProvider != null)
            {
                ClearErrors();
                _taskProvider.Dispose();
                _taskProvider = null;
            }

            if (_designerHost != null)
            {
                _designerHost.Dispose();
                _designerHost = null;
            }

            if (_serviceProvider != null)
            {
                _serviceProvider.Dispose();
                _serviceProvider = null;
            }

            if (_codeBehindCodeGenerator != null)
            {
                _codeBehindCodeGenerator.Close();
                _codeBehindCodeGenerator = null;
            }

            _parser = null;
            _hierarchy = null;
            _appVirtualPath = null;
            _appPhysicalPath = null;
            _codeExtension = null;
            _lastDocument = null;
            _lastDocumentContents = null;

            GC.SuppressFinalize(this);
        }

        /// <summary>
        ///     Gets a VshierarchyItem for the document if possible
        /// </summary>
        protected virtual VsHierarchyItem? GetDocumentItem(string document)
        {
            if (!string.IsNullOrEmpty(document) && _hierarchy != null)
            {
                return VsHierarchyItem.CreateFromMoniker(document, _hierarchy);
            }

            return null;
        }

        /// <summary>
        ///     Gets a VshierarchyItem for the code behind file if possible.
        ///     Will find the file if codeBehindFile/codeBehind are not set using default naming conventions.
        /// </summary>
        protected virtual VsHierarchyItem? GetCodeItem(string document, string codeBehind, string codeBehindFile)
        {
            if (_hierarchy == null)
            {
                return null;
            }

            VsHierarchyItem? itemCode = null;

            // First try the codeBehindFile
            if (!string.IsNullOrEmpty(codeBehindFile))
            {
                itemCode = VsHierarchyItem.CreateFromMoniker(codeBehindFile, _hierarchy);
                if (itemCode != null)
                {
                    return itemCode;
                }
            }

            // Second try the default naming convention
            if (!string.IsNullOrEmpty(document))
            {
                itemCode = VsHierarchyItem.CreateFromMoniker(document + _codeExtension, _hierarchy);
                if (itemCode != null)
                {
                    return itemCode;
                }
            }

            return itemCode;
        }

        /// <summary>
        ///     Generates the .designer file if necessary and possible.
        /// </summary>
        public virtual void UpdateDesignerClass(string document, string codeBehind, string codeBehindFile, bool force, bool create)
        {
            VsHierarchyItem? itemDocument = null;
            VsHierarchyItem? itemCode = null;
            string documentContents = string.Empty;
            bool fSetInProcessUpdate = false;

            try
            {
                if (_fInProcessUpdateDesignerClass)
                {
                    // already in the UpdateDesignerClass call, return now
                    return;
                }
                else
                {
                    fSetInProcessUpdate = true;
                    _fInProcessUpdateDesignerClass = true;
                }

                // Locate the document, codebehind, and .designer files in the hierarchy
                itemDocument = GetDocumentItem(document);
                itemCode = GetCodeItem(document, codeBehind, codeBehindFile);

                if (itemDocument != null && itemCode != null && _appPhysicalPath != null &&
                   _parser != null && (force || itemDocument.IsDirty || !itemDocument.IsRedoStackEmpty()))
                {
                    // Ensure we have moniker path to code behind file
                    codeBehindFile = itemCode.FullPath;

                    if (_codeBehindCodeGenerator != null && _codeBehindCodeGenerator.IsGenerateAllowed(document, codeBehindFile, create))
                    {
                        // Get document contents
                        documentContents = itemDocument.GetDocumentText();

                        if (force || HasDocumentChanged(document, documentContents))
                        {
                            ClearErrors();

                            // Cache last document info
                            _lastDocument = document;
                            _lastDocumentContents = documentContents;

                            // Calculate required params for parse
                            string? virtualPath = CodeGenUtils.MakeRelativeUrl(document, _appPhysicalPath);

                            // Begin the parse
                            _parser.BeginParse(virtualPath, documentContents);

                            // Parse the File
                            _parser.Parse();

                            // Begin generation
                            _codeBehindCodeGenerator.BeginGenerate(document, codeBehindFile, _parser.ClassName_Full, create);

                            // Add the fields
                            if (_parser.ControlInfos != null)
                            {
                                foreach (ControlInfo controlInfo in _parser.ControlInfos)
                                {
                                    EnsureControlDeclaration(controlInfo);
                                }
                            }

                            // Add strong type properties
                            if (!Strings.IsNullOrEmpty(_parser.MasterPageTypeName))
                            {
                                _codeBehindCodeGenerator.EnsureStronglyTypedProperty("Master", _parser.MasterPageTypeName);
                            }
                            if (!Strings.IsNullOrEmpty(_parser.PreviousPageTypeName))
                            {
                                _codeBehindCodeGenerator.EnsureStronglyTypedProperty("PreviousPage", _parser.PreviousPageTypeName);
                            }

                            // Generate Code
                            _codeBehindCodeGenerator.Generate();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMessage = string.Empty;
                try
                {
                    errorMessage = string.Format(WebResources.Generator_GenerationFailed, Path.GetFileName(document), ex.Message);
                }
                catch
                {
                }

                LogError(errorMessage, document, documentContents, ex);

                if (create)
                {
                    // Return error to migration
                    throw new Exception(errorMessage, ex);
                }
                else
                {
                    // Report the error
                    ReportDocumentError(errorMessage, document, documentContents, ex);
                }
            }
            finally
            {
                if (fSetInProcessUpdate)
                {
                    _fInProcessUpdateDesignerClass = false;
                }
            }
        }

        /// <summary>
        ///     Examines the parsed control builder to determine if a control
        ///     declaration should be created.  If so, it calls the generator
        ///     to do generate the field.
        /// </summary>
        private void EnsureControlDeclaration(ControlInfo controlInfo)
        {
            // Ignore it if it doesn't have an ID
            string? id = controlInfo.ID;
            if (Strings.IsNullOrEmpty(id))
            {
                return;
            }

            // Get the control type name
            string? typeName = controlInfo?.DeclareTypeName;
            if (Strings.IsNullOrEmpty(typeName))
            {
                return;
            }

            _codeBehindCodeGenerator?.EnsureControlDeclaration(id, typeName);
        }

        /// <summary>
        ///     Add warning that generation failed for document
        /// </summary>
        private void ReportDocumentError(string errorMessage, string document, string documentContents, Exception ex)
        {
            try
            {
                ClearErrorsForDocument(document);

                ErrorTask task = new ErrorTask(ex);
                task.CanDelete = true;
                task.Document = document;
                task.ErrorCategory = TaskErrorCategory.Warning;
                task.HierarchyItem = _hierarchy;
                task.Text = errorMessage;
                task.Navigate += new EventHandler(OnTaskNavigate);

                if (ex is HttpParseException pex)
                {
                    int line = pex.Line;
                    if (line > 0)
                    {
                        string? file = null;
                        if (pex.VirtualPath != null)
                        {
                            file = pex.FileName;
                        }

                        if (file == null || string.Compare(file, document, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            task.Line = line - Parser.LineOffset;
                        }
                    }
                }

                _taskProvider?.Tasks?.Add(task);
            }
            catch
            {
            }
        }

        /// <summary>
        ///     If logging log the error to the log file
        /// </summary>
        private void LogError(string errorMessage, string document, string? documentContents, Exception ex)
        {
            try
            {
                string? logFile = Environment.GetEnvironmentVariable("WAP_LOG_CODEGENFAILURES");

                if (!Strings.IsNullOrEmpty(logFile))
                {
                    string line = string.Empty;
                    if (ex is HttpParseException pex)
                    {
                        line = ":" + (pex.Line - Parser.LineOffset).ToString(CultureInfo.InvariantCulture);
                    }
                    using (StreamWriter sw = File.AppendText(logFile))
                    {
                        sw.WriteLine("-------------------------------------------------------------");
                        sw.WriteLine(document + line);
                        sw.WriteLine(errorMessage);
                        sw.WriteLine("-------------------------------------------------------------");
                        sw.WriteLine(ex.ToString());
                        sw.WriteLine("-------------------------------------------------------------");
                        if (documentContents != null)
                        {
                            sw.WriteLine(documentContents);
                        }
                        sw.WriteLine("-------------------------------------------------------------");
                    }
                }
            }
            catch
            {
            }
        }

        /// <summary>
        ///     Route navigate request to task provider
        /// </summary>
        private void OnTaskNavigate(object sender, EventArgs e)
        {
            if (_taskProvider != null)
            {
                _taskProvider.Navigate((Task)sender, VSConstants.LOGVIEWID_Code);
            }
        }

        /// <summary>
        ///     Clears out errors from the task provider for the document specified
        /// </summary>
        private void ClearErrorsForDocument(string document)
        {
            if (_taskProvider == null)
            {
                return;
            }
            try
            {
                List<Task> tasks = new List<Task>();

                foreach (Task task in _taskProvider.Tasks)
                {
                    if (string.Compare(task.Document, document, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        tasks.Add(task);
                    }
                }

                foreach (Task task in tasks)
                {
                    _taskProvider.Tasks.Remove(task);
                }
            }
            catch
            {
            }
        }

        /// <summary>
        ///     Clears out errors from the task provider
        /// </summary>
        private void ClearErrors()
        {
            try
            {
                _taskProvider?.Tasks?.Clear();
            }
            catch
            {
            }
        }

        /// <summary>
        ///     Checks in cache to see if document has been modified since last generate
        ///     (we only cache the last document)
        /// </summary>
        protected bool HasDocumentChanged(string document, string documentContents)
        {
            if (!Strings.IsNullOrEmpty(_lastDocument)
                && !Strings.IsNullOrEmpty(_lastDocumentContents)
                && !Strings.IsNullOrEmpty(document)
                && !Strings.IsNullOrEmpty(documentContents)
                && _lastDocument.Length == document.Length
                && _lastDocumentContents.Length == documentContents.Length
                && string.Compare(_lastDocument, document, StringComparison.OrdinalIgnoreCase) == 0
                && string.Compare(_lastDocumentContents, documentContents, StringComparison.Ordinal) == 0)
            {
                return false;
            }

            return true;
        }
    }
}

