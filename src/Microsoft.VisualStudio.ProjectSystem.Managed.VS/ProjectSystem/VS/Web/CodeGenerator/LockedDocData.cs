using System;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Design.Serialization;
using Microsoft.VisualStudio.Shell.Interop;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    /// <summary>
    ///
    /// LockedDocData - A thin wrapper around Microsoft.VisualStudio.Shell.Design.Serialization.DocData
    ///                 that adds an extra edit lock.  The only reason for doing this is so that we can release
    ///                 the last lock when the file was closed and cause a save.
    ///
    /// </summary>
    internal class LockedDocData : DocData
    {
        private readonly RunningDocumentTable _rdt;
        private uint _cookie;

        /// <summary>
        /// Constructor.  Aquires edit lock on document.
        /// </summary>
        public LockedDocData(IServiceProvider serviceProvider, string fileName) : base(serviceProvider, fileName)
        {
            _rdt = new RunningDocumentTable(serviceProvider);

            // Locate and lock the document
            _rdt.FindDocument(fileName, out _cookie);
            _rdt.LockDocument(_VSRDTFLAGS.RDT_EditLock, _cookie);
        }

        /// <summary>
        /// Override of Dispose so we can free our lock after DocData.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                if (_cookie != 0 && _rdt != null)
                {
                    // prevent recursion
                    uint cookie = _cookie;
                    _cookie = 0;

                    try
                    {
                        // Unlock the document, specifying to save if this is the last lock and the buffer is dirty
                        _rdt.UnlockDocument(_VSRDTFLAGS.RDT_EditLock | _VSRDTFLAGS.RDT_Unlock_SaveIfDirty, cookie);
                    }
                    finally
                    {
                        _cookie = 0;
                    }
                }
            }
        }
    }
}
