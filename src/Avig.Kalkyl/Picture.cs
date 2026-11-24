using System.Buffers.Binary;

namespace Avig.Kalkyl;

/// <summary>The image formats a <see cref="Picture"/> can hold.</summary>
public enum PictureFormat
{
    /// <summary>PNG.</summary>
    Png,

    /// <summary>JPEG.</summary>
    Jpeg,

    /// <summary>GIF.</summary>
    Gif,
}

/// <summary>
/// An image to put on a sheet with <see cref="SheetWriter.AddPicture(Picture, CellAddress, double)"/>
/// or <see cref="Worksheet.AddPicture(Picture, CellAddress, double)"/>. PNG, JPEG and GIF are
/// accepted; the format and the size in pixels are read from the file's header, so the image is
/// stored exactly as given.
/// </summary>
public sealed record Picture
{
    private readonly byte[] data;

    private Picture(byte[] data, PictureFormat format, int width, int height)
    {
        this.data = data;
        Format = format;
        Width = width;
        Height = height;
    }

    /// <summary>The format.</summary>
    public PictureFormat Format { get; }

    /// <summary>The width in pixels.</summary>
    public int Width { get; }

    /// <summary>The height in pixels.</summary>
    public int Height { get; }

    /// <summary>The alternative text that screen readers read, or null. Set it with <c>picture with { Description = "..." }</c>.</summary>
    public string? Description { get; init; }

    internal ReadOnlyMemory<byte> Data => data;

    internal string Extension => Format switch
    {
        PictureFormat.Jpeg => "jpeg",
        PictureFormat.Gif => "gif",
        _ => "png",
    };

    internal string ContentType => "image/" + Extension;

    /// <summary>Creates a picture from the bytes of an image file. The bytes are copied.</summary>
    /// <param name="data">The file's bytes.</param>
    /// <returns>The picture, or Invalid when the bytes are not a PNG, JPEG or GIF whose size can be read.</returns>
    public static Result<Picture> FromBytes(ReadOnlySpan<byte> data)
    {
        if (TryReadHeader(data, out PictureFormat format, out int width, out int height) == false)
        {
            return Result<Picture>.Invalid("The image is not a PNG, JPEG or GIF file, or its header is damaged.");
        }

        return new Picture(data.ToArray(), format, width, height);
    }

    /// <summary>Reads a picture from an image file.</summary>
    /// <param name="path">The file.</param>
    /// <param name="cancellationToken">Cancels the read.</param>
    /// <returns>The picture, NotFound when there is no such file, or Invalid when it is not a PNG, JPEG or GIF.</returns>
    public static async Task<Result<Picture>> FromFileAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        byte[] bytes;

        try
        {
            bytes = await File.ReadAllBytesAsync(path, cancellationToken);
        }
        catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException)
        {
            return Result<Picture>.NotFound($"'{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result<Picture>.Forbidden($"'{path}' may not be read.");
        }
        catch (IOException exception)
        {
            return Result<Picture>.Conflict($"'{path}' could not be read: {exception.Message}");
        }

        if (TryReadHeader(bytes, out PictureFormat format, out int width, out int height) == false)
        {
            return Result<Picture>.Invalid($"'{path}' is not a PNG, JPEG or GIF file, or its header is damaged.");
        }

        return new Picture(bytes, format, width, height);
    }

    /// <summary>
    /// Finds the format and size. Every read is checked against the length, since the bytes may
    /// come from anyone: a damaged header gives false, never an exception or a long loop.
    /// </summary>
    internal static bool TryReadHeader(ReadOnlySpan<byte> data, out PictureFormat format, out int width, out int height)
    {
        format = PictureFormat.Png;
        width = 0;
        height = 0;
        bool found = false;

        if (data.Length >= 24 && data[..8].SequenceEqual((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) == true
            && data.Slice(12, 4).SequenceEqual("IHDR"u8) == true)
        {
            format = PictureFormat.Png;
            width = (int)Math.Min(BinaryPrimitives.ReadUInt32BigEndian(data.Slice(16, 4)), int.MaxValue);
            height = (int)Math.Min(BinaryPrimitives.ReadUInt32BigEndian(data.Slice(20, 4)), int.MaxValue);
            found = true;
        }
        else if (data.Length >= 10 && (data[..6].SequenceEqual("GIF87a"u8) == true || data[..6].SequenceEqual("GIF89a"u8) == true))
        {
            format = PictureFormat.Gif;
            width = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(6, 2));
            height = BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(8, 2));
            found = true;
        }
        else if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xD8)
        {
            format = PictureFormat.Jpeg;
            found = TryReadJpegSize(data, out width, out height);
        }

        return found == true && width > 0 && height > 0;
    }

    /// <summary>Walks the JPEG segments to the first frame header (SOF), which holds the size.</summary>
    private static bool TryReadJpegSize(ReadOnlySpan<byte> data, out int width, out int height)
    {
        width = 0;
        height = 0;
        int position = 2;

        // Each step moves forward by at least two bytes, so the loop ends.
        while (position + 4 <= data.Length)
        {
            if (data[position] != 0xFF)
            {
                return false;
            }

            byte marker = data[position + 1];

            // Fill bytes between segments.
            if (marker == 0xFF)
            {
                position++;

                continue;
            }

            // Markers without a length: TEM and RST0 to RST7.
            if (marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7))
            {
                position += 2;

                continue;
            }

            // The end of the image, or the start of the compressed data, before any frame header.
            if (marker == 0xD9 || marker == 0xDA)
            {
                return false;
            }

            int length = BinaryPrimitives.ReadUInt16BigEndian(data.Slice(position + 2, 2));

            if (length < 2)
            {
                return false;
            }

            // SOF0 to SOF15, except DHT (C4), JPG (C8) and DAC (CC), which share the range.
            bool frame = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;

            if (frame == true)
            {
                if (position + 9 > data.Length)
                {
                    return false;
                }

                height = BinaryPrimitives.ReadUInt16BigEndian(data.Slice(position + 5, 2));
                width = BinaryPrimitives.ReadUInt16BigEndian(data.Slice(position + 7, 2));

                return true;
            }

            position += 2 + length;
        }

        return false;
    }
}
