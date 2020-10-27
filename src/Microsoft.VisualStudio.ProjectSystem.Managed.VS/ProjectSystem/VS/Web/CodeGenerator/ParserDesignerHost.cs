using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Reflection;
using System.Web.UI;
using System.Web.UI.Design;
using Microsoft.VisualStudio.Shell.Design;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Web.Application;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    /// <summary>
    ///
    ///     ParserDesignerHost - A simple designer host that exposes the client type resolution
    ///                          service to the field genration parser.
    ///
    /// </summary>
    internal class ParserDesignerHost : IDesignerHost, ITypeResolutionService, IUserControlTypeResolutionService
    {
        private IServiceProvider? _serviceProvider;
        private IVsHierarchy? _hierarchy;
        private ITypeResolutionService? _typeResolutionService;
        private Parser? _parser;

        /// <summary>
        ///     Constructor 
        /// </summary>
        public ParserDesignerHost(IServiceProvider serviceProvider, IVsHierarchy hierarchy)
        {
            _serviceProvider = serviceProvider;
            _hierarchy = hierarchy;
            _parser = null;
        }

        /// <summary>
        ///     Finalizer
        /// </summary>
        ~ParserDesignerHost()
        {
            System.Diagnostics.Debug.Fail("ParserDesignerHost was not disposed.");
            Dispose();
        }

        /// <summary>
        ///     Cleanup
        /// </summary>
        public void Dispose()
        {
            _serviceProvider = null;
            _hierarchy = null;
            _typeResolutionService = null;
            _parser = null;
            GC.SuppressFinalize(this);
        }

        /// <summary>
        ///     Fetch the client ITypeResolution service for the project
        /// </summary>
        private ITypeResolutionService? TypeResolutionService
        {
            get
            {
                if (_typeResolutionService == null)
                {
                    if (_serviceProvider != null)
                    {
                        if (_serviceProvider.GetService(typeof(DynamicTypeService)) is DynamicTypeService dynamicTypeService)
                        {
                            _typeResolutionService = dynamicTypeService.GetTypeResolutionService(_hierarchy);
                        }
                    }
                }
                return _typeResolutionService;
            }
        }

        /// <summary>
        ///     Set circular ref to parser.
        /// </summary>
        public void SetParser(Parser parser)
        {
            _parser = parser;
        }

        /// <summary>
        ///     Block/unblock the generator from item type resolution
        ///     The genrator was designed to only resolve types that are referenced by going through the 
        ///     VSTypeResolutionService.  However, when an editor is open the Venus VsItemTypeResolutionService is active
        ///     and app domain assembly resolve calls will be sent to it.  This function alows the generator to turn on/off
        ///     all item type resolvers in the web application.  We make sure to turn it off beform making any calls to the 
        ///     VSTypeResolutionService and turn it back on afterwards.
        /// </summary>
        public void BlockItemTypeResolution(bool block)
        {
            if (_hierarchy is IVsWebApplicationProject waProject)
            {
                waProject.BlockItemTypeResolver(block);
            }
        }

        /// <summary>
        ///     Returns ourseves as the type resolution service
        ///     we then delegate to the client project ITypeResolutionService
        ///     for most calls.
        /// </summary>
        object? IServiceProvider.GetService(Type serviceType)
        {
            if (typeof(ITypeResolutionService).IsEquivalentTo(serviceType))
            {
                if (TypeResolutionService != null)
                {
                    return this;
                }
            }
            else if (typeof(IUserControlTypeResolutionService).IsEquivalentTo(serviceType))
            {
                if (_parser != null && TypeResolutionService != null)
                {
                    return this;
                }
            }
            else if (typeof(IWebApplication).IsEquivalentTo(serviceType))
            {
                if (_hierarchy != null)
                {
                    WAProject proj = WAProject.GetProjectFromIVsHierarchy(_hierarchy);
                    if (proj != null)
                    {
                        return proj.GetContextService<IWebApplication>();
                    }
                }
            }
            else if (typeof(IVsHierarchy).IsEquivalentTo(serviceType))
            {
                return _hierarchy;
            }

            return null;
        }

        /// <summary>
        ///     Activate (not implemented)
        /// </summary>
        void IDesignerHost.Activate()
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.Activate");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     Activated (not implemented)
        /// </summary>
        event EventHandler IDesignerHost.Activated
        {
            add
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.Activated");
                throw new NotImplementedException();
            }
            remove
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.Activated");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     Container (not implemented)
        /// </summary>
        IContainer IDesignerHost.Container
        {
            get
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.Container");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     CreateComponent (not implemented)
        /// </summary>
        IComponent IDesignerHost.CreateComponent(Type componentClass, string name)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.CreateComponent");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     CreateComponent (not implemented)
        /// </summary>
        IComponent IDesignerHost.CreateComponent(Type componentClass)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.CreateComponent");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     CreateTransaction (not implemented)
        /// </summary>
        DesignerTransaction IDesignerHost.CreateTransaction(string description)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.CreateTransaction");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     CreateTransaction (not implemented)
        /// </summary>
        DesignerTransaction IDesignerHost.CreateTransaction()
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.CreateTransaction");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     Deactivated (not implemented)
        /// </summary>
        event EventHandler IDesignerHost.Deactivated
        {
            add
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.Deactivated");
                throw new NotImplementedException();
            }
            remove
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.Deactivated");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     DestroyComponent (not implemented)
        /// </summary>
        void IDesignerHost.DestroyComponent(IComponent component)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.DestroyComponent");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     GetDesigner (not implemented)
        /// </summary>
        IDesigner IDesignerHost.GetDesigner(IComponent component)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.GetDesigner");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     GetType delegate to client type resolution
        /// </summary>
        Type? IDesignerHost.GetType(string typeName)
        {
            BlockItemTypeResolution(true);
            try
            {
                return TypeResolutionService?.GetType(typeName, false, true);
            }
            finally
            {
                BlockItemTypeResolution(false);
            }
        }

        /// <summary>
        ///     InTransaction (not implemented)
        /// </summary>
        bool IDesignerHost.InTransaction
        {
            get
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.InTransaction");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     LoadComplete (not implemented)
        /// </summary>
        event EventHandler IDesignerHost.LoadComplete
        {
            add
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.LoadComplete");
                throw new NotImplementedException();
            }
            remove
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.LoadComplete");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     Loading (not implemented)
        /// </summary>
        bool IDesignerHost.Loading
        {
            get
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.Loading");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     We currently return null for the RootComponent when requested.
        ///     This seems to be good enough for generation.
        /// </summary>
        IComponent? IDesignerHost.RootComponent
        {
            get
            {
                // Generation may request a root component

                // One case of this is when NamespaceTagNameToTypeMapper looks for the WebFormsReferenceManager
                // to call its version of GetType.  If we return null it just uses the regular type resolutions service.

                return null;
            }
        }

        /// <summary>
        ///     RootComponentClassName (not implemented)
        /// </summary>
        string IDesignerHost.RootComponentClassName
        {
            get
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.RootComponentClassName");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     TransactionClosed (not implemented)
        /// </summary>
        event DesignerTransactionCloseEventHandler IDesignerHost.TransactionClosed
        {
            add
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionClosed");
                throw new NotImplementedException();
            }
            remove
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionClosed");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     TransactionClosing (not implemented)
        /// </summary>
        event DesignerTransactionCloseEventHandler IDesignerHost.TransactionClosing
        {
            add
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionClosing");
                throw new NotImplementedException();
            }
            remove
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionClosing");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     TransactionDescription (not implemented)
        /// </summary>
        string IDesignerHost.TransactionDescription
        {
            get
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionDescription");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     TransactionOpened (not implemented)
        /// </summary>
        event EventHandler IDesignerHost.TransactionOpened
        {
            add
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionOpened");
                throw new NotImplementedException();
            }
            remove
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionOpened");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     TransactionOpening (not implemented)
        /// </summary>
        event EventHandler IDesignerHost.TransactionOpening
        {
            add
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionOpening");
                throw new NotImplementedException();
            }
            remove
            {
                System.Diagnostics.Debug.Fail("Unexpected call to IDesignerHost.TransactionOpening");
                throw new NotImplementedException();
            }
        }

        /// <summary>
        ///     AddService (not implemented)
        /// </summary>
        void IServiceContainer.AddService(Type serviceType, ServiceCreatorCallback callback, bool promote)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IServiceContainer.AddService");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     AddService (not implemented)
        /// </summary>
        void IServiceContainer.AddService(Type serviceType, ServiceCreatorCallback callback)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IServiceContainer.AddService");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     AddService (not implemented)
        /// </summary>
        void IServiceContainer.AddService(Type serviceType, object serviceInstance, bool promote)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IServiceContainer.AddService");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     AddService (not implemented)
        /// </summary>
        void IServiceContainer.AddService(Type serviceType, object serviceInstance)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IServiceContainer.AddService");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     RemoveService (not implemented)
        /// </summary>
        void IServiceContainer.RemoveService(Type serviceType, bool promote)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IServiceContainer.RemoveService");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     RemoveService (not implemented)
        /// </summary>
        void IServiceContainer.RemoveService(Type serviceType)
        {
            System.Diagnostics.Debug.Fail("Unexpected call to IServiceContainer.RemoveService");
            throw new NotImplementedException();
        }

        /// <summary>
        ///     GetAssembly delegates to client type resolution service
        /// </summary>
        Assembly? ITypeResolutionService.GetAssembly(AssemblyName name, bool throwOnError)
        {
            BlockItemTypeResolution(true);
            try
            {
                return TypeResolutionService?.GetAssembly(name, throwOnError);
            }
            finally
            {
                BlockItemTypeResolution(false);
            }
        }

        /// <summary>
        ///     GetAssembly delegates to client type resolution service
        /// </summary>
        Assembly? ITypeResolutionService.GetAssembly(AssemblyName name)
        {
            BlockItemTypeResolution(true);
            try
            {
                return TypeResolutionService?.GetAssembly(name);
            }
            finally
            {
                BlockItemTypeResolution(false);
            }
        }

        /// <summary>
        ///     GetPathOfAssembly delegates to client type resolution service
        /// </summary>
        string? ITypeResolutionService.GetPathOfAssembly(AssemblyName name)
        {
            BlockItemTypeResolution(true);
            try
            {
                return TypeResolutionService?.GetPathOfAssembly(name);
            }
            finally
            {
                BlockItemTypeResolution(false);
            }
        }

        /// <summary>
        ///     GetType delegates to client type resolution service
        /// </summary>
        Type? ITypeResolutionService.GetType(string name, bool throwOnError, bool ignoreCase)
        {
            BlockItemTypeResolution(true);
            try
            {
                return TypeResolutionService?.GetType(name, throwOnError, ignoreCase);
            }
            finally
            {
                BlockItemTypeResolution(false);
            }
        }

        /// <summary>
        ///     GetType delegates to client type resolution service
        /// </summary>
        Type? ITypeResolutionService.GetType(string name, bool throwOnError)
        {
            BlockItemTypeResolution(true);
            try
            {
                return TypeResolutionService?.GetType(name, throwOnError);
            }
            finally
            {
                BlockItemTypeResolution(false);
            }
        }

        /// <summary>
        ///     GetType delegates to client type resolution service
        /// </summary>
        Type? ITypeResolutionService.GetType(string name)
        {
            BlockItemTypeResolution(true);
            try
            {
                return TypeResolutionService?.GetType(name);
            }
            finally
            {
                BlockItemTypeResolution(false);
            }
        }

        /// <summary>
        ///     ReferenceAssembly does nothing.
        ///     We do not want to delegate to the client type resolution service because
        ///     it could add a reference and cause checkouts.  Generation should never
        ///     add references.
        /// </summary>
        void ITypeResolutionService.ReferenceAssembly(AssemblyName name)
        {
            // The generator is not allowed to add references
        }

        /// <summary>
        ///     Gets the base type of the user control.  We do not get the actual type because it not likely available
        ///     at design time.
        /// </summary>
        Type? IUserControlTypeResolutionService.GetType(string tagPrefix, string tagName)
        {
            BlockItemTypeResolution(true);
            try
            {
                if (_parser != null && TypeResolutionService != null)
                {
                    string typeName = _parser.GetUserControlTypeName(tagPrefix, tagName);
                    return TypeResolutionService.GetType(typeName, true);
                }

                return null;
            }
            finally
            {
                BlockItemTypeResolution(false);
            }
        }
    }
}
