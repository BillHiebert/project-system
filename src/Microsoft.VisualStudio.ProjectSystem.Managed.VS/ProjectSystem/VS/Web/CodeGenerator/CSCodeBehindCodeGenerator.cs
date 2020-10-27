// Licensed to the .NET Foundation under one or more agreements. The .NET Foundation licenses this file to you under the MIT license. See the LICENSE.md file in the project root for more information.

using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    [Guid("349c585a-65df-11da-9384-00065b846f21")]
    internal class CSWACodeBehindCodeGenerator : WACodeBehindCodeGenerator
    {
        /// <summary>
        /// C# Override to not declare form when the ID is the same as the class
        /// as this causes a compile error and we did not do it at all in VS 2003.
        /// </summary>
        protected override bool ShouldDeclareField(string name, string typeName)
        {
            bool declareField = base.ShouldDeclareField(name, typeName);

            if (declareField)
            {
                string? className_Name = ClassName_Name;

                // Don't add field for Form if ID is same as class
                if (!Strings.IsNullOrEmpty(className_Name)
                    && string.Compare(typeName, "System.Web.UI.HtmlControls.HtmlForm", StringComparison.OrdinalIgnoreCase) == 0
                    && string.Compare(name, className_Name, StringComparison.Ordinal) == 0)
                {
                    declareField = false;
                }
            }

            return declareField;
        }

        /// <summary>
        /// C# override to generate code
        /// </summary>
        protected override string GenerateCode()
        {
            if (CodeDomProvider == null || CodeDomProvider.FileExtension != "cs")
            {
                return base.GenerateCode();
            }

            StringWriter stringWriter = new StringWriter(CultureInfo.InvariantCulture);

            stringWriter.WriteLine("//------------------------------------------------------------------------------");
            stringWriter.Write("// <");
            stringWriter.Write(WebResources.Generator_AutoGen_Comment_Tag);
            stringWriter.WriteLine(">");
            stringWriter.Write("//     ");
            stringWriter.WriteLine(WebResources.Generator_AutoGen_Comment_Line1);
            stringWriter.WriteLine("//");
            stringWriter.Write("//     ");
            stringWriter.WriteLine(WebResources.Generator_AutoGen_Comment_Line2);
            stringWriter.Write("//     ");
            stringWriter.WriteLine(WebResources.Generator_AutoGen_Comment_Line3);
            stringWriter.Write("// </");
            stringWriter.Write(WebResources.Generator_AutoGen_Comment_Tag);
            stringWriter.WriteLine(">");
            stringWriter.WriteLine("//------------------------------------------------------------------------------");
            stringWriter.WriteLine("");

            CodeDomProvider.GenerateCodeFromNamespace(CodeNamespace, stringWriter, CodeGeneratorOptions);

            return stringWriter.ToString();
        }

        protected override SyntaxTree GetSyntaxTree(string generatedCode)
        {
            return CSharpSyntaxTree.ParseText(generatedCode);
        }
    }
}
