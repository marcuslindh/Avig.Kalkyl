using System.Xml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// An XmlReader that refuses elements nested deeper than <see cref="MaxDepth"/>. Spreadsheet
/// XML never goes past about ten levels, but 200 000 nested &lt;r&gt; in one cell (1.4 MB of
/// file) took 28 s and 107 GB of allocations in the document model, because copying a cell's
/// XML costs time in proportion to its depth at every level. The error is an XmlException,
/// which every caller already turns into a damaged file.
/// </summary>
internal sealed class DepthLimitedXmlReader : XmlReader, IXmlLineInfo
{
    /// <summary>Ten times what Excel writes, and far below where depth starts to cost.</summary>
    public const int MaxDepth = 100;

    private readonly XmlReader inner;

    /// <summary>
    /// True when the inner reader is synchronous. Its async methods then answer from the
    /// synchronous ones, so code written against the async API can read an in-memory fragment
    /// with a synchronous reader, whose buffer is 4 kB rather than the 64 kB an async reader takes.
    /// </summary>
    private readonly bool synchronous;

    private readonly XmlReaderSettings? settings;

    public DepthLimitedXmlReader(XmlReader inner)
    {
        this.inner = inner;
        settings = inner.Settings;
        synchronous = inner.Settings?.Async != true;

        if (synchronous == true && settings is not null)
        {
            settings = settings.Clone();
            settings.Async = true;
        }
    }

    public override int AttributeCount => inner.AttributeCount;

    public override string BaseURI => inner.BaseURI;

    public override int Depth => inner.Depth;

    public override bool EOF => inner.EOF;

    public override bool HasValue => inner.HasValue;

    public override bool IsDefault => inner.IsDefault;

    public override bool IsEmptyElement => inner.IsEmptyElement;

    public override string LocalName => inner.LocalName;

    public override string Name => inner.Name;

    public override string NamespaceURI => inner.NamespaceURI;

    public override XmlNameTable NameTable => inner.NameTable;

    public override XmlNodeType NodeType => inner.NodeType;

    public override string Prefix => inner.Prefix;

    public override ReadState ReadState => inner.ReadState;

    public override XmlReaderSettings? Settings => settings;

    public override string Value => inner.Value;

    public override string XmlLang => inner.XmlLang;

    public override XmlSpace XmlSpace => inner.XmlSpace;

    /// <summary>The line in the part, for error messages; 0 when the inner reader has none.</summary>
    public int LineNumber
    {
        get
        {
            if (inner is IXmlLineInfo info)
            {
                return info.LineNumber;
            }

            return 0;
        }
    }

    public int LinePosition
    {
        get
        {
            if (inner is IXmlLineInfo info)
            {
                return info.LinePosition;
            }

            return 0;
        }
    }

    public bool HasLineInfo()
    {
        return inner is IXmlLineInfo info && info.HasLineInfo();
    }

    public override string GetAttribute(int index)
    {
        return inner.GetAttribute(index);
    }

    public override string? GetAttribute(string name)
    {
        return inner.GetAttribute(name);
    }

    public override string? GetAttribute(string name, string? namespaceURI)
    {
        return inner.GetAttribute(name, namespaceURI);
    }

    public override string? LookupNamespace(string prefix)
    {
        return inner.LookupNamespace(prefix);
    }

    public override bool MoveToAttribute(string name)
    {
        return inner.MoveToAttribute(name);
    }

    public override bool MoveToAttribute(string name, string? ns)
    {
        return inner.MoveToAttribute(name, ns);
    }

    public override void MoveToAttribute(int index)
    {
        inner.MoveToAttribute(index);
    }

    public override bool MoveToElement()
    {
        return inner.MoveToElement();
    }

    public override bool MoveToFirstAttribute()
    {
        return inner.MoveToFirstAttribute();
    }

    public override bool MoveToNextAttribute()
    {
        return inner.MoveToNextAttribute();
    }

    public override bool Read()
    {
        bool more = inner.Read();
        Check();

        return more;
    }

    public override async Task<bool> ReadAsync()
    {
        if (synchronous == true)
        {
            return Read();
        }

        bool more = await inner.ReadAsync();
        Check();

        return more;
    }

    public override Task<string> GetValueAsync()
    {
        if (synchronous == true)
        {
            return Task.FromResult(inner.Value);
        }

        return inner.GetValueAsync();
    }

    // Straight to the inner reader for speed: an element read as text has no child elements, so
    // there is no depth to check. Skip stays on this reader: a deep subtree skipped when the file
    // is opened would otherwise be found only when the file is saved and the subtree copied.
    public override string ReadElementContentAsString()
    {
        return inner.ReadElementContentAsString();
    }

    public override Task<string> ReadElementContentAsStringAsync()
    {
        if (synchronous == true)
        {
            return Task.FromResult(inner.ReadElementContentAsString());
        }

        return inner.ReadElementContentAsStringAsync();
    }

    public override bool ReadAttributeValue()
    {
        return inner.ReadAttributeValue();
    }

    public override void ResolveEntity()
    {
        inner.ResolveEntity();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing == true)
        {
            inner.Dispose();
        }

        base.Dispose(disposing);
    }

    private void Check()
    {
        if (inner.Depth > MaxDepth)
        {
            throw new XmlException($"Elements are nested more than {MaxDepth} levels deep, which no spreadsheet program writes.", null, LineNumber, LinePosition);
        }
    }
}
