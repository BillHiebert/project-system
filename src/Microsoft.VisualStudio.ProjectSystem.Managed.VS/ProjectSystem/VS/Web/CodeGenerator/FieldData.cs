using System;
using System.Collections.Generic;
using EnvDTE;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    /// <summary>
    ///
    /// FieldData - encapsulates properties of fields to allow generator to
    ///             decide how/if it should generate fields in the .designer file
    ///
    /// </summary>
    internal class FieldData
    {
        public FieldData(CodeClass codeClass, CodeVariable codeVariable, int depth)
        {
            Name = codeVariable.Name;
            Class = codeClass.FullName;
            Depth = depth;
            vsCMAccess access = codeVariable.Access;
            IsPrivate = ((access & vsCMAccess.vsCMAccessPrivate) == vsCMAccess.vsCMAccessPrivate);
            IsProtected = ((access & vsCMAccess.vsCMAccessProtected) == vsCMAccess.vsCMAccessProtected);

            // Special casing vsCMAccessAssemblyOrFamily
            if ((access & vsCMAccess.vsCMAccessAssemblyOrFamily) == vsCMAccess.vsCMAccessAssemblyOrFamily)
            {
                IsProtected = true;
            }

            IsPublic = ((access & vsCMAccess.vsCMAccessPublic) == vsCMAccess.vsCMAccessPublic);
        }

        public string Class { get; private set; }
        public string Name { get; private set; }
        public bool IsPrivate { get; private set; }
        public bool IsProtected { get; private set; }
        public bool IsPublic { get; private set; }
        public int Depth { get; private set; }
    }

    internal class FieldDataDictionary : Dictionary<string, FieldData>
    {
        public FieldDataDictionary(bool caseSensitive) : base(caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase)
        {
        }
    }
}
