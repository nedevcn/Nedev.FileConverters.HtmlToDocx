using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Net;

namespace Nedev.FileConverters.HtmlToDocx.Core.Html;

public ref struct HtmlParser
{
    private ReadOnlySpan<char> _input;
    private int _position;

    public HtmlParser(ReadOnlySpan<char> input)
    {
        _input = input;
        _position = 0;
    }

    public HtmlToken ParseNext()
    {
        SkipWhitespace();
        if (_position >= _input.Length) return HtmlToken.EofToken();
        if (_input[_position] == '<') return ParseTag();
        return ParseText();
    }

    public HtmlToken[] ParseAll()
    {
        var tokens = new List<HtmlToken>();
        HtmlToken token;
        while ((token = ParseNext()).Type != HtmlTokenType.EOF)
            tokens.Add(token);
        return tokens.ToArray();
    }

    private HtmlToken ParseTag()
    {
        _position++; // skip '<'
        if (_position >= _input.Length) return HtmlToken.TextToken("<");

        if (_input[_position] == '!')
        {
            if (_position + 3 <= _input.Length && _input.Slice(_position, 3).Equals("!--", StringComparison.Ordinal))
            {
                _position += 3;
                int start = _position;
                int end = _input.Slice(_position).IndexOf("-->", StringComparison.Ordinal);
                if (end == -1)
                {
                    _position = _input.Length;
                    return HtmlToken.CommentToken(_input.Slice(start).ToString());
                }
                string comment = _input.Slice(start, end).ToString();
                _position += end + 3;
                return HtmlToken.CommentToken(comment);
            }
            
            if (_position + 8 <= _input.Length && _input.Slice(_position, 8).Equals("!doctype", StringComparison.OrdinalIgnoreCase))
            {
                int start = _position;
                while (_position < _input.Length && _input[_position] != '>') _position++;
                string doctype = _input.Slice(start, _position - start).ToString();
                if (_position < _input.Length) _position++;
                return HtmlToken.DoctypeToken(doctype);
            }
        }

        bool isEndTag = _input[_position] == '/';
        if (isEndTag) _position++;

        int nameStart = _position;
        while (_position < _input.Length && !char.IsWhiteSpace(_input[_position]) && 
               _input[_position] != '>' && _input[_position] != '/')
            _position++;

        if (nameStart == _position) return HtmlToken.TextToken("<");
        string tagName = _input.Slice(nameStart, _position - nameStart).ToString().ToLowerInvariant();
        var attributes = ParseAttributes();

        bool isSelfClosing = false;
        if (_position < _input.Length && _input[_position] == '/')
        {
            isSelfClosing = true;
            _position++;
        }
        if (_position < _input.Length && _input[_position] == '>') _position++;

        if (isEndTag) return HtmlToken.EndTagToken(tagName);
        if (isSelfClosing) return HtmlToken.SelfClosingTagToken(tagName, attributes);
        return HtmlToken.StartTagToken(tagName, attributes);
    }

    private HtmlAttribute[] ParseAttributes()
    {
        var attributes = new List<HtmlAttribute>();
        while (_position < _input.Length)
        {
            SkipWhitespace();
            if (_position >= _input.Length || _input[_position] == '>' || _input[_position] == '/') break;

            int nameStart = _position;
            while (_position < _input.Length && !char.IsWhiteSpace(_input[_position]) && 
                   _input[_position] != '=' && _input[_position] != '>' && _input[_position] != '/')
                _position++;

            if (nameStart == _position) break;
            string name = _input.Slice(nameStart, _position - nameStart).ToString().ToLowerInvariant();
            string? value = null;

            if (_position < _input.Length && _input[_position] == '=')
            {
                _position++;
                value = ParseAttributeValue();
            }
            attributes.Add(new HtmlAttribute(name, value));
        }
        return attributes.ToArray();
    }

    private string? ParseAttributeValue()
    {
        SkipWhitespace();
        if (_position >= _input.Length) return null;

        char quote = _input[_position];
        if (quote == '"' || quote == '\'')
        {
            _position++;
            int valueStart = _position;
            while (_position < _input.Length && _input[_position] != quote) _position++;
            string value = _input.Slice(valueStart, _position - valueStart).ToString();
            if (_position < _input.Length && _input[_position] == quote) _position++;
            return WebUtility.HtmlDecode(value);
        }

        int unquotedStart = _position;
        while (_position < _input.Length && !char.IsWhiteSpace(_input[_position]) && _input[_position] != '>')
            _position++;
        return WebUtility.HtmlDecode(_input.Slice(unquotedStart, _position - unquotedStart).ToString());
    }

    private HtmlToken ParseText()
    {
        int start = _position;
        while (_position < _input.Length && _input[_position] != '<') _position++;
        return HtmlToken.TextToken(WebUtility.HtmlDecode(_input.Slice(start, _position - start).ToString()));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void SkipWhitespace()
    {
        while (_position < _input.Length && char.IsWhiteSpace(_input[_position])) _position++;
    }
}
