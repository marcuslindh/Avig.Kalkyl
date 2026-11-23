namespace Avig.Kalkyl;

/// <summary>
/// Marks a partial class or record as a row type. A source generator then writes the code that
/// turns instances into rows and rows into instances, so no reflection is used at run time.
/// </summary>
/// <example>
/// <code>
/// [KalkylRow]
/// public partial class Customer
/// {
///     [KalkylColumn("Name", Width = 30)]
///     public required string Name { get; init; }
///
///     [KalkylColumn(NumberFormat = "# ##0.00")]
///     public decimal Balance { get; init; }
///
///     public DateOnly? Since { get; init; }
/// }
/// </code>
/// </example>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
public sealed class KalkylRowAttribute : Attribute
{
}

/// <summary>Sets how a property is written as a column and which header it is read from.</summary>
[AttributeUsage(AttributeTargets.Property, Inherited = false)]
public sealed class KalkylColumnAttribute : Attribute
{
    /// <summary>Uses the property name as the header.</summary>
    public KalkylColumnAttribute()
    {
    }

    /// <summary>Uses another header than the property name.</summary>
    /// <param name="name">The header.</param>
    public KalkylColumnAttribute(string name)
    {
        Name = name;
    }

    /// <summary>The header, or null for the property name. Headers are matched without regard to case or surrounding spaces.</summary>
    public string? Name { get; }

    /// <summary>The column's position when writing; lower comes first. Properties without one follow in declaration order.</summary>
    public int Order { get; set; } = int.MaxValue;

    /// <summary>The number format when writing, such as <c>0.00</c> or <c>yyyy-mm-dd</c>.</summary>
    public string? NumberFormat { get; set; }

    /// <summary>The column width when writing, in characters. 0 leaves the default.</summary>
    public double Width { get; set; }
}

/// <summary>Leaves a property out of the mapping.</summary>
[AttributeUsage(AttributeTargets.Property, Inherited = false)]
public sealed class KalkylIgnoreAttribute : Attribute
{
}
