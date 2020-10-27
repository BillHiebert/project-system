namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    // Simple parser to parse the directive from page/usercontrol/masterpage content
    internal class DirectiveParser : ASPParser
    {
        public DirectiveParser(string text)
        {
            BeginParse(text);
        }

        public Directive? ParseNextDirective()
        {
            ASPElement? element = ParseElement(ASPElementType.Directive);

            if (element != null)
            {
                return new Directive(element);
            }

            return null;
        }
    }
}
