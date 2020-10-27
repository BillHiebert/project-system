using System.Collections.Generic;
using System.Text;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    // Simple struct to store information about a directive.
    internal class Directive
    {
        private string? _name;
        private bool _isMain;
        private readonly Dictionary<string, string> _attributes = new Dictionary<string, string>();

        public Directive()
        {
        }

        public Directive(ASPElement element)
        {
            if (element.Type == ASPElementType.Directive)
            {
                _name = element.LCaseName;

                if (string.IsNullOrEmpty(_name) || _name == "page" || _name == "control" || _name == "master")
                {
                    _isMain = true;
                }
                if (element.Attributes != null)
                {
                    foreach (ASPAttribute attribute in element.Attributes)
                    {
                        if (attribute.LCaseName != null && attribute.Value != null)
                        {
                            _attributes[attribute.LCaseName] = attribute.Value;
                        }
                    }
                }
            }
        }

        internal string? Name
        {
            get { return _name; }
            set { _name = value; }
        }

        internal Dictionary<string, string> Attributes
        {
            get { return _attributes; }
        }

        internal bool IsMain
        {
            get { return _isMain || _name == null; }
            set { _isMain = value; }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();

            sb.Append("<%@ ");
            sb.Append(Name);

            foreach (KeyValuePair<string, string> entry in Attributes)
            {
                string key = entry.Key;
                if (!string.IsNullOrEmpty(key))
                {
                    sb.Append(" ");
                    sb.Append(key);
                    sb.Append("=\"");

                    string val = entry.Value;
                    if (!string.IsNullOrEmpty(val))
                    {
                        sb.Append(val);
                    }

                    sb.Append("\"");
                }
            }

            sb.Append(" %>");

            return sb.ToString();
        }
    }

    internal class DirectiveList : List<Directive>
    {
        public DirectiveList()
        {
        }

        public override string ToString()
        {
            var sb = new StringBuilder();

            foreach (Directive directive in this)
            {
                sb.Append(directive.ToString());
                sb.Append("\r\n");
            }

            return sb.ToString();
        }
    }
}
