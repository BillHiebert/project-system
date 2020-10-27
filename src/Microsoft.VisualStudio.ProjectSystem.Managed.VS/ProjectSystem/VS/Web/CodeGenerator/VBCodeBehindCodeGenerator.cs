using System;
using System.CodeDom;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    [Guid("349c585b-65df-11da-9384-00065b846f21")]
    internal class VBWACodeBehindCodeGenerator : WACodeBehindCodeGenerator
    {
        private string? _implicitNamespace;

        /// <summary>
        /// VB Override to clean up additional state
        /// </summary>
        protected override void DisposeGenerateState()
        {
            base.DisposeGenerateState();
            _implicitNamespace = null;
        }

        /// <summary>
        /// Calculate implicit namespace of code behind
        /// </summary>
        protected string? ImplicitNamespace
        {
            get
            {
                if (_implicitNamespace == null)
                {
                    if (ItemCode != null)
                    {
                        try
                        {
                            _implicitNamespace = ItemCode.DefaultNamespace;
                        }
                        catch
                        {
                        }
                    }
                }

                return _implicitNamespace;
            }
        }

        /// <summary>
        /// VB Override to add option strict and option explicit
        /// </summary>
        protected override CodeCompileUnit GenerateCodeCompileUnit()
        {
            CodeCompileUnit ccu = base.GenerateCodeCompileUnit();

            // Add stringent VB compile settings
            ccu.UserData.Add("AllowLateBound", false);
            ccu.UserData.Add("RequireVariableDeclaration", true);

            return ccu;
        }

        /// <summary>
        /// VB Override to adjust namespace generated so that it does not include the implicit part
        /// </summary>
        protected override string? GetClassNamespace()
        {
            string? className_Namespace = ClassName_Namespace;
            string? className_Full = ClassName_Full;
            string? implicitNamespace = ImplicitNamespace;

            // Strip off implicit namespace if any
            if (!Strings.IsNullOrEmpty(className_Namespace)
                && !Strings.IsNullOrEmpty(implicitNamespace)
                && !Strings.IsNullOrEmpty(className_Full)
                && className_Full.StartsWith(implicitNamespace + ".", StringComparison.Ordinal))
            {
                if (className_Namespace.Length > implicitNamespace.Length)
                {
                    className_Namespace = className_Namespace.Substring(implicitNamespace.Length + 1);
                }
                else
                {
                    className_Namespace = null;
                }
            }

            return className_Namespace;
        }

        /// <summary>
        /// VB Override to enable WithEvents on declarations
        /// </summary>
        protected override void SetAdditionalFieldData(CodeMemberField field)
        {
            base.SetAdditionalFieldData(field);

            // Enable WithEvents or controls in VB
            field.UserData["WithEvents"] = true;
        }

        /// <summary>
        /// VB override to generate code
        /// </summary>
        protected override string GenerateCode()
        {
            if (CodeDomProvider == null || CodeDomProvider.FileExtension != "vb")
            {
                return base.GenerateCode();
            }

            StringWriter stringWriter = new StringWriter(CultureInfo.InvariantCulture);

            stringWriter.WriteLine("'------------------------------------------------------------------------------");
            stringWriter.Write("' <");
            stringWriter.Write(WebResources.Generator_AutoGen_Comment_Tag);
            stringWriter.WriteLine(">");
            stringWriter.Write("'     ");
            stringWriter.WriteLine(WebResources.Generator_AutoGen_Comment_Line1);
            stringWriter.WriteLine("'");
            stringWriter.Write("'     ");
            stringWriter.WriteLine(WebResources.Generator_AutoGen_Comment_Line2);
            stringWriter.Write("'     ");
            stringWriter.WriteLine(WebResources.Generator_AutoGen_Comment_Line3);
            stringWriter.Write("' </");
            stringWriter.Write(WebResources.Generator_AutoGen_Comment_Tag);
            stringWriter.WriteLine(">");
            stringWriter.WriteLine("'------------------------------------------------------------------------------");
            stringWriter.WriteLine("");

            stringWriter.WriteLine("Option Strict On");
            stringWriter.WriteLine("Option Explicit On");

            CodeDomProvider.GenerateCodeFromNamespace(CodeNamespace, stringWriter, CodeGeneratorOptions);

            return stringWriter.ToString();
        }

        protected override SyntaxTree GetSyntaxTree(string generatedCode)
        {
            return VisualBasicSyntaxTree.ParseText(generatedCode);
        }
    }
}
