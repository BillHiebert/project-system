
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.CodeDom;
using System.ComponentModel.Design;
using System.Web;
using System.Web.Configuration;
using System.Web.Management;
using System.Web.RegularExpressions;
using System.Web.UI;
using System.Web.UI.Design;
using System.Web.UI.WebControls;
using System.Net;

namespace Microsoft.VisualStudio.ProjectSystem.VS.Web
{
    internal class ASPParser
    {
        internal static readonly Regex _tagRegex40 = new TagRegex();
        internal static readonly Regex _tagRegex35 = new TagRegex35();
        internal static readonly Regex _directiveRegex = new DirectiveRegex();
        internal static readonly Regex _endtagRegex = new EndTagRegex();
        internal static readonly Regex _aspCodeRegex = new AspCodeRegex();
        internal static readonly Regex _aspExprRegex = new AspExprRegex();
        internal static readonly Regex _databindExprRegex = new DatabindExprRegex();
        internal static readonly Regex _commentRegex = new CommentRegex();
        internal static readonly Regex _includeRegex = new IncludeRegex();
        internal static readonly Regex _textRegex = new TextRegex();
        internal static readonly Version VERSION40 = new Version(4, 0);

        private string? _text;
        private int _textPos;
        private Match? _match;
        private ASPElement? _element;
        private ASPTextPositionList? _positions;
        private bool _inServerScript;
        private int _lastGTIndex;
        private bool _unrecognizedTag;
        private Version? _version;
        private Regex? _tagRegex;

        public void Dispose()
        {
            _text = null;
            _textPos = 0;
            _match = null;
            _element = null;
            _positions = null;
            _inServerScript = false;
            _lastGTIndex = 0;
            _unrecognizedTag = false;
            _version = null;
            _tagRegex = null;
        }

        public void BeginParse(string text)
        {
            BeginParse(null, text);
        }

        public void BeginParse(Version? version, string text)
        {
            _text = text;
            _textPos = 0;
            _lastGTIndex = 0;
            _inServerScript = false;
            _unrecognizedTag = false;
            _positions = new ASPTextPositionList();
            _version = version ?? VERSION40;
            _tagRegex = _version >= VERSION40 ? _tagRegex40 : _tagRegex35;

            OnBeginParse();
        }

        public virtual void OnBeginParse()
        {
        }

        private int LastGTIndex
        {
            get
            {
                // Find the last '>' in the input string
                if (_lastGTIndex == 0 && _text != null)
                {
                    _lastGTIndex = _text.LastIndexOf('>');
                }

                return _lastGTIndex;
            }
        }

        // Parse an element and return it
        public ASPElement? ParseElement()
        {
            return ParseElement(out bool _);
        }

        // Parse an element and return it
        // Also return the OnParsed flag indicating if parsing should continue
        private ASPElement? ParseElement(out bool continueParsing)
        {
            continueParsing = false;
            _element = null;
            if (_text == null)
            {
                return _element;
            }

            while (_element == null && _textPos < _text.Length)
            {
                if (ParseText()                      // ...<  or ...eoi
                    || ParseDirective()              // <%@ ... %>
                    || ParseInclude()                // <!-- #incude ... --> 
                    || ParseComment()                // <%-- ... --%>
                    || ParseCodeExpression()         // <%= ... %>
                    || ParseDataBindingExpression()  // <%# ... %>
                    || ParseCode()                   // <% ... %>
                    || ParseTag()                    // < ... >
                    || ParseEndTag())                // </ ... >
                {
                }
                else
                {
                    // tag format was not recognized
                    // so parse up to next tag as text
                    _unrecognizedTag = true;
                }
            }

            if (_element != null)
            {
                continueParsing = OnParsed(_element);
            }

            return _element;
        }

        // Parse until the specified element type is found and return it
        public ASPElement? ParseElement(ASPElementType elementType)
        {
            ASPElement? element = ParseElement();

            while (element != null && element.Type != elementType)
            {
                element = ParseElement();
            }

            return element;
        }

        public virtual bool OnParsed(ASPElement element)
        {
            return true;
        }

        // Literal Text
        protected bool ParseText()
        {
            if (_text == null)
            {
                return false;
            }

            // determine if we should fail or include '<'
            int offset = 0; // scan current
            if (_text[_textPos] == '<')
            {
                if (_unrecognizedTag)
                {
                    offset = 1; // include current '<' as text
                    _unrecognizedTag = false;
                }
                else
                {
                    offset = -1; // don't parse as text 
                }
            }

            if (offset >= 0 && (_match = _textRegex.Match(_text, _textPos + offset)).Success)
            {
                _element = CreateElement(ASPElementType.Text, CreateSpan(_match.Index - offset, _match.Length + offset));
                _textPos = _match.Index + _match.Length;

                return true;
            }
            else if (offset > 0)
            {
                // If there is not a text match and we are at an unrecognized tag then return the "<" as text
                _element = CreateElement(ASPElementType.Text, CreateSpan(_textPos, offset));
                _textPos = _textPos + offset;

                return true;
            }

            return false;
        }

        // Directive Block
        // <%@ ... %>
        protected bool ParseDirective()
        {
            if (!_inServerScript && _text != null && (_match = _directiveRegex.Match(_text, _textPos)).Success)
            {
                ASPTextSpan span = CreateSpan(_match.Index, _match.Length);
                ASPTextSpan? name = null;
                ASPAttributeList? attributes = null;

                CaptureCollection attrNames = _match.Groups["attrname"].Captures;
                CaptureCollection attrValues = _match.Groups["attrval"].Captures;

                for (int i = 0; i < attrNames.Count; i++)
                {
                    Capture captureName = attrNames[i];
                    //Capture captureEqual = attrEquals[i];
                    Capture captureValue = attrValues[i];

                    ASPTextSpan spanName = CreateSpan(captureName.Index, captureName.Length);

                    ASPTextSpan? spanValue = null;
                    char beforeValue = _text[captureValue.Index - 1];
                    char afterValue = _text[captureValue.Index + captureValue.Length];
                    if ((beforeValue == '\"' && afterValue == '\"') || (beforeValue == '\'' && afterValue == '\''))
                    {
                        spanValue = CreateSpan(captureValue.Index - 1, captureValue.Length + 2);
                    }
                    else if (captureValue.Length > 0)
                    {
                        spanValue = CreateSpan(captureValue.Index, captureValue.Length);
                    }

                    if (i == 0 && spanValue == null)
                    {
                        name = spanName;
                    }
                    else
                    {
                        if (attributes == null)
                        {
                            attributes = new ASPAttributeList();
                        }
                        ASPAttribute attribute = new ASPAttribute(spanName, spanValue);
                        attributes.Add(attribute);
                    }
                }

                _element = CreateElement(ASPElementType.Directive, span, name, attributes);
                _textPos = _match.Index + _match.Length;

                return true;
            }

            return false;
        }

        // Server Side Include
        // <!-- #include file="foo.inc" -->
        protected bool ParseInclude()
        {
            if (!_inServerScript && (_match = _includeRegex.Match(_text, _textPos)).Success)
            {
                _element = CreateElement(ASPElementType.Include, CreateSpan(_match.Index, _match.Length));
                _textPos = _match.Index + _match.Length;

                return true;
            }

            return false;
        }

        // Comment Block
        // <%-- ... --%> 
        protected bool ParseComment()
        {
            if (!_inServerScript && (_match = _commentRegex.Match(_text, _textPos)).Success)
            {
                _element = CreateElement(ASPElementType.Comment, CreateSpan(_match.Index, _match.Length));
                _textPos = _match.Index + _match.Length;

                return true;
            }

            return false;
        }

        // Expression Code Block
        // <%= ... %>
        protected bool ParseCodeExpression()
        {
            if (!_inServerScript && (_match = _aspExprRegex.Match(_text, _textPos)).Success)
            {
                _element = CreateElement(ASPElementType.CodeExpression, CreateSpan(_match.Index, _match.Length));
                _textPos = _match.Index + _match.Length;

                return true;
            }

            return false;
        }

        // Databinding Expression Block
        // <%# ... %>
        // (this does not include blocks used as values for attributes of server tags)
        protected bool ParseDataBindingExpression()
        {
            if (!_inServerScript && (_match = _databindExprRegex.Match(_text, _textPos)).Success)
            {
                _element = CreateElement(ASPElementType.DataBindingExpression, CreateSpan(_match.Index, _match.Length));
                _textPos = _match.Index + _match.Length;

                return true;
            }

            return false;
        }

        // Code Block
        // <% ... %>
        protected bool ParseCode()
        {
            if (!_inServerScript && (_match = _aspCodeRegex.Match(_text, _textPos)).Success)
            {
                _element = CreateElement(ASPElementType.Code, CreateSpan(_match.Index, _match.Length));
                _textPos = _match.Index + _match.Length;

                return true;
            }

            return false;
        }

        protected bool ParseTag()
        {
            if (_text != null && !_inServerScript && _tagRegex != null && (LastGTIndex > _textPos) && (_match = _tagRegex.Match(_text, _textPos)).Success)
            {
                ASPTextSpan span = CreateSpan(_match.Index, _match.Length);
                ASPAttributeList? attributes = null;
                bool isRunAtServer = false;
                bool hasID = false;

                Capture captureTagName = _match.Groups["tagname"];
                ASPTextSpan name = CreateSpan(captureTagName.Index, captureTagName.Length);
                bool fSelfClosed = _match.Groups["empty"].Success;

                CaptureCollection attrNames = _match.Groups["attrname"].Captures;
                CaptureCollection attrValues = _match.Groups["attrval"].Captures;

                for (int i = 0; i < attrNames.Count; i++)
                {
                    Capture captureName = attrNames[i];
                    Capture captureValue = attrValues[i];

                    ASPTextSpan spanName = CreateSpan(captureName.Index, captureName.Length);

                    ASPTextSpan? spanValue = null;
                    char beforeValue = _text[captureValue.Index - 1];
                    char afterValue = _text[captureValue.Index + captureValue.Length];
                    if ((beforeValue == '\"' && afterValue == '\"') || (beforeValue == '\'' && afterValue == '\''))
                    {
                        spanValue = CreateSpan(captureValue.Index - 1, captureValue.Length + 2);
                    }
                    else if (captureValue.Length > 0)
                    {
                        spanValue = CreateSpan(captureValue.Index, captureValue.Length);
                    }

                    if (attributes == null)
                    {
                        attributes = new ASPAttributeList();
                    }
                    ASPAttribute attribute = new ASPAttribute(spanName, spanValue);
                    attributes.Add(attribute);

                    string? lcaseAttrName = attribute.LCaseName;
                    if (lcaseAttrName == "runat")
                    {
                        string? lcaseAttrValue = attribute.LCaseValue;
                        if (lcaseAttrValue == "server")
                        {
                            isRunAtServer = true;
                        }
                    }
                    else if (lcaseAttrName == "id")
                    {
                        hasID = true;
                    }
                }

                _element = CreateElement(ASPElementType.Tag, span, name, attributes, isRunAtServer, hasID, fSelfClosed);
                _textPos = _match.Index + _match.Length;

                if (_element.IsRunAtServer && _element.LCaseName == "script")
                {
                    _inServerScript = true;
                }

                return true;
            }

            return false;
        }

        protected bool ParseEndTag()
        {
            if ((_match = _endtagRegex.Match(_text, _textPos)).Success)
            {
                ASPTextSpan span = CreateSpan(_match.Index, _match.Length);

                Capture captureTagName = _match.Groups["tagname"];
                ASPTextSpan name = CreateSpan(captureTagName.Index, captureTagName.Length);

                _element = CreateElement(ASPElementType.EndTag, span, name, null);
                _textPos = _match.Index + _match.Length;

                if (_inServerScript && _element.LCaseName == "script")
                {
                    _inServerScript = false;
                }

                return true;
            }

            return false;
        }

        internal ASPTextSpan? Insert(int position, string text)
        {
            if (string.IsNullOrEmpty(text) || Strings.IsNullOrEmpty(_text))
            {
                return null;
            }

            // get the length of the text to be inserted
            int length = text.Length;

            // insert the text
            _text = _text.Insert(position, text);

            // update the text position markers
            if (_positions != null)
            {
                foreach (ASPTextPosition textPosition in _positions)
                {
                    if (textPosition != null && textPosition.Offset >= position)
                    {
                        textPosition.Offset += length;
                    }
                }
            }

            // create the span for the new text
            ASPTextSpan? span = CreateSpan(position, length);

            return span;
        }

        protected ASPElement CreateElement(ASPElementType elementType)
        {
            return CreateElement(elementType, null, null, null, false, false, false);
        }

        protected ASPElement CreateElement(ASPElementType elementType, ASPTextSpan span)
        {
            return CreateElement(elementType, span, null, null, false, false, false);
        }

        protected ASPElement CreateElement(ASPElementType elementType, ASPTextSpan span, ASPTextSpan? name, ASPAttributeList? attributes)
        {
            return CreateElement(elementType, span, name, attributes, false, false, false);
        }

        protected ASPElement CreateElement(ASPElementType elementType, ASPTextSpan span, ASPTextSpan? name, ASPAttributeList? attributes, bool isRunAtServer, bool hasID)
        {
            return CreateElement(elementType, span, name, attributes, isRunAtServer, hasID, false);
        }

        protected virtual ASPElement CreateElement(ASPElementType elementType, ASPTextSpan? span, ASPTextSpan? name, ASPAttributeList? attributes, bool isRunAtServer, bool hasID, bool isSelfClosed)
        {
            return new ASPElement(elementType, span, name, attributes, isRunAtServer, hasID, isSelfClosed);
        }

        private ASPTextSpan CreateSpan(int position, int length)
        {
            ASPTextPosition start = new ASPTextPosition(position);
            ASPTextPosition end = new ASPTextPosition(position + length - 1);

            if (_positions != null)
            {
                _positions.Add(start);
                _positions.Add(end);
            }

            ASPTextSpan span = new ASPTextSpan(start, end, this);
            return span;
        }

        public string GetText(ASPTextPosition start, ASPTextPosition end)
        {
            return _text == null ? string.Empty : _text.Substring(start.Offset, end.Offset - start.Offset + 1);
        }

        public string Text
        {
            get
            {
                return _text == null? string.Empty : _text;
            }
        }
    }

    internal enum ASPElementType
    {
        Document,
        Text,
        Directive,
        Include,
        Comment,
        CodeExpression,
        DataBindingExpression,
        Code,
        Tag,
        EndTag,
    }

    internal class ASPElement
    {
        private ASPElementType _type;
        private readonly ASPTextSpan? _span;
        private readonly ASPTextSpan? _name;
        private readonly ASPAttributeList? _attributes;
        private readonly bool _isRunAtServer;
        private readonly bool _hasID;
        private readonly bool _isSelfClosed;

        public ASPElement(ASPElementType type, ASPTextSpan? span, ASPTextSpan? name, ASPAttributeList? attributes, bool isRunAtServer, bool hasID, bool isSelfClosed)
        {
            _type = type;
            _span = span;
            _name = name;
            _attributes = attributes;
            _isRunAtServer = isRunAtServer;
            _hasID = hasID;
            _isSelfClosed = isSelfClosed;
        }

        public ASPElementType Type
        {
            get
            {
                return _type;
            }
        }

        public ASPTextSpan? Span
        {
            get
            {
                return _span;
            }
        }

        public ASPParser? Parser
        {
            get
            {
                return _span?.Parser;
            }
        }

        public ASPTextSpan? Name
        {
            get
            {
                return _name;
            }
        }

        public ASPAttributeList? Attributes
        {
            get
            {
                return _attributes;
            }
        }

        public bool IsRunAtServer
        {
            get
            {
                return _isRunAtServer;
            }
        }

        public bool HasID
        {
            get
            {
                return _hasID;
            }
        }

        public bool IsSelfClosed
        {
            get
            {
                return _isSelfClosed;
            }
        }

        public override string ToString()
        {
            if (Span != null)
            {
                return Span.ToString();
            }
            return string.Empty;
        }

        public string? LCaseName
        {
            get
            {
                if (Name != null)
                {
                    return Name.ToString().ToLower(CultureInfo.InvariantCulture);
                }

                return null;
            }
        }

        public string? LCaseNamespace
        {
            get
            {
                string? lcName = LCaseName;
                if (!Strings.IsNullOrEmpty(lcName))
                {
                    int firstColon = lcName.IndexOf(':');
                    if (firstColon > 0)
                    {
                        return lcName.Substring(0, firstColon).Trim();
                    }
                }

                return null;
            }
        }

        public string? GetAttributeValue(string lcasename)
        {
            if (Attributes != null)
            {
                foreach (ASPAttribute attribute in Attributes)
                {
                    if (attribute.LCaseName == lcasename)
                    {
                        return attribute.Value;
                    }
                }
            }

            return null;
        }

        public ASPAttribute? AddAttribute(string name, string value)
        {
            return AddAttribute(name, value, false, true);
        }

        public ASPAttribute? AddAttribute(string name, string value, bool htmlEncode, bool quoted)
        {
            // This is not general purpose yet
            // We only handle adding attributes to directives
            if (_type != ASPElementType.Directive)
            {
                return null;
            }

            // must provide name, or if we don't have a span
            if (string.IsNullOrEmpty(name) || Span == null || Parser == null)
            {
                return null;
            }

            ASPTextSpan? valueSpan = null;
            ASPTextSpan? nameSpan;
            int position;

            // the insertion position is the end of the directive on the %
            position = Span.End.Offset - 1;

            // if necessary insert leading space
            if (!char.IsWhiteSpace(Parser.Text[position - 1]))
            {
                Parser.Insert(position, " ");
                position += 1;
            }

            // insert name
            nameSpan = Parser.Insert(position, name);
            if (nameSpan != null)
            {
                position += nameSpan.Length;
            }

            if (value != null)
            {
                // insert =
                Parser.Insert(position, "=");
                position += 1;

                if (htmlEncode && !string.IsNullOrEmpty(value))
                {
                    value = WebUtility.HtmlEncode(value);
                }

                if (quoted)
                {
                    value = "\"" + value + "\"";
                }

                // insert value
                valueSpan = Parser.Insert(position, value);
                position += value.Length;
            }

            // insert trailing space
            Parser.Insert(position, " ");
            position += 1;

            // Add the attribute
            ASPAttribute? attr = null;
            if (nameSpan != null && valueSpan != null)
            {
                attr = new ASPAttribute(nameSpan, valueSpan);
                _attributes?.Add(attr);
            }

            return attr;
        }
    }

    internal class ASPTextPosition
    {
        public int Offset;

        public ASPTextPosition(int offset)
        {
            Offset = offset;
        }
    }

    internal class ASPTextSpan
    {
        private readonly ASPTextPosition _start;
        private readonly ASPTextPosition _end;
        private readonly ASPParser _parser;

        public ASPTextSpan(ASPTextPosition start, ASPTextPosition end, ASPParser parser)
        {
            _start = start;
            _end = end;
            _parser = parser;
        }

        public override string ToString()
        {
            if (_parser != null && _start != null && _end != null)
            {
                return _parser.GetText(_start, _end);

            }
            return string.Empty;
        }

        public ASPTextPosition Start
        {
            get
            {
                return _start;
            }
        }

        public ASPTextPosition End
        {
            get
            {
                return _end;
            }
        }

        public int Length
        {
            get
            {
                return _end.Offset - _start.Offset + 1;
            }
        }

        public ASPParser Parser
        {
            get
            {
                return _parser;
            }
        }
    }

    internal class ASPAttribute
    {
        private readonly ASPTextSpan? _nameSpan;
        private readonly ASPTextSpan? _valueSpan;

        public ASPAttribute(ASPTextSpan? nameSpan, ASPTextSpan? valueSpan)
        {
            _nameSpan = nameSpan;
            _valueSpan = valueSpan;
        }

        public override string ToString()
        {
            if (_nameSpan != null)
            {
                string text = _nameSpan.ToString();
                if (_valueSpan != null)
                {
                    text = text + "=" + _valueSpan.ToString();
                }
                return text;
            }
            return string.Empty;
        }

        public string? Name
        {
            get
            {
                if (_nameSpan != null)
                {
                    return _nameSpan.ToString();
                }

                return null;
            }
        }

        public string? LCaseName
        {
            get
            {
                string? name = Name;
                if (name != null)
                {
                    return name.ToLower(CultureInfo.InvariantCulture);
                }

                return null;
            }
        }

        public string? RawValue
        {
            get
            {
                if (_valueSpan != null)
                {
                    return _valueSpan.ToString();
                }

                return null;
            }
        }

        public string? DecodedRawValue
        {
            get
            {
                string? val = RawValue;
                if (val != null)
                {
                    return WebUtility.HtmlDecode(val);
                }

                return null;
            }
        }

        // returns the unquoted and decoded value
        public string? Value
        {
            get
            {
                string? val = DecodedRawValue;
                if (val != null)
                {
                    // remove quotes
                    if ((val.StartsWith("\"", StringComparison.Ordinal) && val.EndsWith("\"", StringComparison.Ordinal))
                        || (val.StartsWith("'", StringComparison.Ordinal) && val.EndsWith("'", StringComparison.Ordinal)))
                    {
                        val = val.Substring(1, val.Length - 2);
                    }

                    return val;
                }

                return null;
            }
        }

        public string? LCaseValue
        {
            get
            {
                string? val = Value;
                if (val != null)
                {
                    return val.ToLower(CultureInfo.InvariantCulture);
                }

                return null;
            }
        }

    }

    internal class ASPAttributeList : List<ASPAttribute>
    {
    }

    internal class ASPTextPositionList : List<ASPTextPosition>
    {
    }

    internal class ASPTreeElementList : List<ASPTreeElement>
    {
    }

    internal class ASPTreeElement : ASPElement
    {
        public ASPTreeElement? Parent;
        public ASPTreeElementList Children = new ASPTreeElementList();
        public bool IsClosed;
        public ASPTreeElement? End;
        public ASPTextSpan? Outer;

        public ASPTreeElement(ASPElementType type, ASPTextSpan? span, ASPTextSpan? name, ASPAttributeList? attributes, bool isRunAtServer, bool hasID, bool isSelfClosed)
            : base(type, span, name, attributes, isRunAtServer, hasID, isSelfClosed)
        {
            IsClosed = isSelfClosed;
            Outer = span;
        }
    }

    internal class ASPTreeParser : ASPParser
    {
        private ASPTreeElement? _root;

        public ASPTreeParser()
        {
        }

        public ASPTreeElement? Root
        {
            get
            {
                return _root;
            }
        }

        public override void OnBeginParse()
        {
            _root = (ASPTreeElement)CreateElement(ASPElementType.Document);
        }

        public override bool OnParsed(ASPElement aspElement)
        {
            if (Root == null || Root.Children == null)
            {
                System.Diagnostics.Debug.Fail("Root element or its children is null");
                return true;
            }

            ASPTreeElement element = (ASPTreeElement)aspElement;
            ASPTreeElement parent = Root;
            ASPTreeElementList elements = Root.Children;

            // If we have an end tag
            if (element.Type == ASPElementType.EndTag)
            {
                // Scan backwards
                int cElements = elements.Count;
                for (int iElement = cElements - 1; iElement >= 0; iElement--)
                {
                    ASPTreeElement current = elements[iElement];

                    // Until we find an unclosed start tag
                    if (current.Type == ASPElementType.Tag
                        && !current.IsClosed)
                    {
                        if (current.LCaseName == element.LCaseName && current.Span != null && element.Span != null)
                        {
                            // close the element
                            current.IsClosed = true;
                            current.End = element;
                            current.Outer = new ASPTextSpan(current.Span.Start, element.Span.End, this);

                            // Move the inner elements to be children
                            int iInner = iElement + 1;
                            int cInner = cElements - iInner;
                            if (cInner > 0)
                            {
                                List<ASPTreeElement> children = elements.GetRange(iInner, cInner);
                                foreach (ASPTreeElement child in children)
                                {
                                    child.Parent = current;
                                }
                                current.Children.AddRange(children);
                                elements.RemoveRange(iInner, cInner);
                            }

                            // stop scanning
                            break;
                        }
                        else if (current.IsRunAtServer && !string.IsNullOrEmpty(current.LCaseNamespace) && string.IsNullOrEmpty(element.LCaseNamespace))
                        {
                            // don't allow simple tag to overlap unclosed namespaced server tag
                            break;
                        }
                    }
                }
            }

            // Add the element
            element.Parent = parent;
            elements?.Add(element);

            return true;
        }

        protected override ASPElement CreateElement(ASPElementType elementType, ASPTextSpan? span, ASPTextSpan? name, ASPAttributeList? attributes, bool isRunAtServer, bool hasID, bool isSelfClosing)
        {
            return new ASPTreeElement(elementType, span, name, attributes, isRunAtServer, hasID, isSelfClosing);
        }
    }
}
