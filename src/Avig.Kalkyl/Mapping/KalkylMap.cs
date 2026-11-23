using System.ComponentModel;

namespace Avig.Kalkyl;

/// <summary>
/// Implemented by the source generator on every <see cref="KalkylRowAttribute"/> type. Lets
/// generic code find the type's mapping without reflection.
/// </summary>
/// <typeparam name="TSelf">The row type.</typeparam>
public interface IKalkylRow<TSelf>
    where TSelf : IKalkylRow<TSelf>
{
    /// <summary>The generated mapping.</summary>
    static abstract KalkylMap<TSelf> KalkylMap { get; }
}

/// <summary>A column in a mapping.</summary>
/// <param name="Header">The header written and matched when reading.</param>
/// <param name="Required">True when a row without a value in the column cannot become an instance.</param>
/// <param name="NumberFormat">The number format when writing, or null.</param>
/// <param name="Width">The column width when writing, or 0 for the default.</param>
public sealed record KalkylColumnInfo(string Header, bool Required, string? NumberFormat, double Width);

/// <summary>
/// The mapping between a row type and the cells of a row. Written by the source generator;
/// you do not normally derive from it yourself.
/// </summary>
/// <typeparam name="T">The row type.</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class KalkylMap<T>
{
    /// <summary>Creates the mapping.</summary>
    /// <param name="columns">The columns, in the order they are written.</param>
    protected KalkylMap(IReadOnlyList<KalkylColumnInfo> columns)
    {
        Columns = columns;
    }

    /// <summary>The columns, in the order they are written.</summary>
    public IReadOnlyList<KalkylColumnInfo> Columns { get; }

    /// <summary>Turns an instance into cells, one per column in <see cref="Columns"/>.</summary>
    /// <param name="item">The instance.</param>
    /// <returns>The cells.</returns>
    public abstract IReadOnlyList<Cell> ToCells(T item);

    /// <summary>Turns a row into an instance.</summary>
    /// <param name="row">The row, with the columns already matched to headers.</param>
    /// <returns>The instance, or Invalid with the cell that could not be read.</returns>
    public abstract Result<T> FromRow(MappedRow row);
}
