using System.Buffers.Binary;
using System.IO.Compression;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// Makes Excel's PNG export smaller. Excel writes 8-bit RGBA without
/// row filters; a picture of a sheet is opaque and has few colours, so it is written again as
/// a palette image when it has at most 256 colours, or as RGB otherwise, with the best filter per
/// row and the strongest compression. Only plain 8-bit RGBA without interlacing is rewritten;
/// anything else is left as it is.
/// </summary>
public static class PngOptimizer
{
    private static readonly byte[] Signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    /// <summary>The size a README picture should stay under; above it the colours are reduced.</summary>
    public const int TargetBytes = 300 * 1024;

    /// <summary>
    /// Rewrites the file in place when that makes it smaller. When the lossless result is still
    /// over <see cref="TargetBytes"/>, the picture is reduced to 256 colours, which text and flat
    /// cards survive well.
    /// </summary>
    public static void Optimize(string path)
    {
        byte[] original = File.ReadAllBytes(path);

        if (Decode(original, out int width, out int height, out byte[] pixels) == false)
        {
            return;
        }

        byte[] smaller = Encode(width, height, pixels);

        if (smaller.Length > TargetBytes)
        {
            Quantize(pixels);
            smaller = Encode(width, height, pixels);
        }

        if (smaller.Length < original.Length)
        {
            File.WriteAllBytes(path, smaller);
        }
    }

    /// <summary>
    /// Reduces the pixels to at most 256 colours: the most common shades (5 bits per channel)
    /// become the palette, each the average of the colours that fall in it, and every pixel takes
    /// the nearest palette colour.
    /// </summary>
    private static void Quantize(byte[] rgba)
    {
        Dictionary<int, (long Count, long Red, long Green, long Blue)> buckets = [];

        for (int index = 0; index < rgba.Length; index += 4)
        {
            int key = ((rgba[index] >> 3) << 10) | ((rgba[index + 1] >> 3) << 5) | (rgba[index + 2] >> 3);
            buckets.TryGetValue(key, out (long Count, long Red, long Green, long Blue) bucket);
            buckets[key] = (bucket.Count + 1, bucket.Red + rgba[index], bucket.Green + rgba[index + 1], bucket.Blue + rgba[index + 2]);
        }

        List<(int Red, int Green, int Blue, long Count)> shades = buckets.Values
            .OrderByDescending(bucket => bucket.Count)
            .Select(bucket => ((int)(bucket.Red / bucket.Count), (int)(bucket.Green / bucket.Count), (int)(bucket.Blue / bucket.Count), bucket.Count))
            .ToList();

        // The most common shades first; then, so that a small icon or a red figure keeps its
        // colour, the shades farthest from the palette so far, weighted by how often they occur.
        List<(int Red, int Green, int Blue)> palette = shades.Take(160).Select(shade => (shade.Red, shade.Green, shade.Blue)).ToList();
        List<(int Red, int Green, int Blue, long Count)> candidates = shades.Skip(160).Where(shade => shade.Count >= 6).ToList();
        double[] farthest = new double[candidates.Count];

        for (int index = 0; index < candidates.Count; index++)
        {
            farthest[index] = double.MaxValue;

            foreach ((int red, int green, int blue) in palette)
            {
                farthest[index] = Math.Min(farthest[index], Distance(candidates[index], red, green, blue));
            }
        }

        while (palette.Count < 256 && candidates.Count > 0)
        {
            int best = 0;
            double bestScore = -1;

            for (int index = 0; index < candidates.Count; index++)
            {
                double score = farthest[index] * Math.Sqrt(candidates[index].Count);

                if (score > bestScore)
                {
                    bestScore = score;
                    best = index;
                }
            }

            (int Red, int Green, int Blue, long Count) added = candidates[best];
            palette.Add((added.Red, added.Green, added.Blue));

            for (int index = 0; index < candidates.Count; index++)
            {
                farthest[index] = Math.Min(farthest[index], Distance(candidates[index], added.Red, added.Green, added.Blue));
            }
        }

        Dictionary<int, int> nearest = [];

        for (int index = 0; index < rgba.Length; index += 4)
        {
            int color = (rgba[index] << 16) | (rgba[index + 1] << 8) | rgba[index + 2];

            if (nearest.TryGetValue(color, out int chosen) == false)
            {
                long best = long.MaxValue;

                for (int slot = 0; slot < palette.Count; slot++)
                {
                    long red = palette[slot].Red - rgba[index];
                    long green = palette[slot].Green - rgba[index + 1];
                    long blue = palette[slot].Blue - rgba[index + 2];
                    long distance = (red * red * 3) + (green * green * 4) + (blue * blue * 2);

                    if (distance < best)
                    {
                        best = distance;
                        chosen = slot;
                    }
                }

                nearest[color] = chosen;
            }

            rgba[index] = (byte)palette[chosen].Red;
            rgba[index + 1] = (byte)palette[chosen].Green;
            rgba[index + 2] = (byte)palette[chosen].Blue;
        }
    }

    private static double Distance((int Red, int Green, int Blue, long Count) shade, int red, int green, int blue)
    {
        double dr = shade.Red - red;
        double dg = shade.Green - green;
        double db = shade.Blue - blue;

        return (dr * dr * 3) + (dg * dg * 4) + (db * db * 2);
    }

    /// <summary>Reads an 8-bit RGBA PNG into rows of RGBA bytes.</summary>
    private static bool Decode(byte[] png, out int width, out int height, out byte[] pixels)
    {
        width = 0;
        height = 0;
        pixels = [];

        if (png.AsSpan().StartsWith(Signature) == false)
        {
            return false;
        }

        using MemoryStream compressed = new();
        int offset = Signature.Length;

        while (offset + 8 <= png.Length)
        {
            int length = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(offset));
            string type = System.Text.Encoding.ASCII.GetString(png, offset + 4, 4);
            ReadOnlySpan<byte> data = png.AsSpan(offset + 8, length);

            if (type == "IHDR")
            {
                width = BinaryPrimitives.ReadInt32BigEndian(data);
                height = BinaryPrimitives.ReadInt32BigEndian(data[4..]);

                // Bit depth 8, colour type 6 (RGBA), no interlacing.
                if (data[8] != 8 || data[9] != 6 || data[12] != 0)
                {
                    return false;
                }
            }
            else if (type == "IDAT")
            {
                compressed.Write(data);
            }

            offset += length + 12;
        }

        compressed.Position = 0;
        using ZLibStream inflater = new(compressed, CompressionMode.Decompress);
        using MemoryStream raw = new();
        inflater.CopyTo(raw);
        byte[] filtered = raw.ToArray();
        int stride = width * 4;

        if (width == 0 || filtered.Length != (stride + 1) * height)
        {
            return false;
        }

        pixels = new byte[stride * height];

        for (int row = 0; row < height; row++)
        {
            byte filter = filtered[row * (stride + 1)];
            Span<byte> line = pixels.AsSpan(row * stride, stride);
            filtered.AsSpan((row * (stride + 1)) + 1, stride).CopyTo(line);
            Span<byte> previous = Span<byte>.Empty;

            if (row > 0)
            {
                previous = pixels.AsSpan((row - 1) * stride, stride);
            }

            Unfilter(filter, line, previous, 4);
        }

        return true;
    }

    private static void Unfilter(byte filter, Span<byte> line, ReadOnlySpan<byte> previous, int bytesPerPixel)
    {
        for (int index = 0; index < line.Length; index++)
        {
            int left = 0;
            int up = 0;
            int upLeft = 0;

            if (index >= bytesPerPixel)
            {
                left = line[index - bytesPerPixel];
            }

            if (previous.IsEmpty == false)
            {
                up = previous[index];

                if (index >= bytesPerPixel)
                {
                    upLeft = previous[index - bytesPerPixel];
                }
            }

            int predicted = filter switch
            {
                1 => left,
                2 => up,
                3 => (left + up) / 2,
                4 => Paeth(left, up, upLeft),
                _ => 0,
            };

            line[index] = (byte)(line[index] + predicted);
        }
    }

    private static int Paeth(int left, int up, int upLeft)
    {
        int estimate = left + up - upLeft;
        int toLeft = Math.Abs(estimate - left);
        int toUp = Math.Abs(estimate - up);
        int toUpLeft = Math.Abs(estimate - upLeft);

        if (toLeft <= toUp && toLeft <= toUpLeft)
        {
            return left;
        }

        if (toUp <= toUpLeft)
        {
            return up;
        }

        return upLeft;
    }

    /// <summary>Writes the pixels as a palette PNG when they have at most 256 colours, as RGB otherwise.</summary>
    private static byte[] Encode(int width, int height, byte[] rgba)
    {
        Dictionary<int, int> palette = [];
        bool fits = true;

        for (int index = 0; index < rgba.Length; index += 4)
        {
            int color = (rgba[index] << 16) | (rgba[index + 1] << 8) | rgba[index + 2];

            if (palette.ContainsKey(color) == false)
            {
                if (palette.Count == 256)
                {
                    fits = false;

                    break;
                }

                palette[color] = palette.Count;
            }
        }

        int bytesPerPixel = 3;
        byte colorType = 2;
        byte[] packed;

        if (fits == true)
        {
            bytesPerPixel = 1;
            colorType = 3;
            packed = new byte[width * height];

            for (int index = 0; index < packed.Length; index++)
            {
                int offset = index * 4;
                packed[index] = (byte)palette[(rgba[offset] << 16) | (rgba[offset + 1] << 8) | rgba[offset + 2]];
            }
        }
        else
        {
            packed = new byte[width * height * 3];

            for (int index = 0; index < width * height; index++)
            {
                packed[index * 3] = rgba[index * 4];
                packed[(index * 3) + 1] = rgba[(index * 4) + 1];
                packed[(index * 3) + 2] = rgba[(index * 4) + 2];
            }
        }

        byte[] filtered = Filter(width, height, packed, bytesPerPixel, fits == false);
        using MemoryStream output = new();
        output.Write(Signature);
        byte[] header = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(header, width);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), height);
        header[8] = 8;
        header[9] = colorType;
        WriteChunk(output, "IHDR", header);

        if (fits == true)
        {
            byte[] entries = new byte[palette.Count * 3];

            foreach ((int color, int slot) in palette)
            {
                entries[slot * 3] = (byte)(color >> 16);
                entries[(slot * 3) + 1] = (byte)(color >> 8);
                entries[(slot * 3) + 2] = (byte)color;
            }

            WriteChunk(output, "PLTE", entries);
        }

        using MemoryStream compressed = new();

        using (ZLibStream deflater = new(compressed, CompressionLevel.SmallestSize, true))
        {
            deflater.Write(filtered);
        }

        WriteChunk(output, "IDAT", compressed.ToArray());
        WriteChunk(output, "IEND", []);

        return output.ToArray();
    }

    /// <summary>
    /// Filters each row. A palette image compresses best unfiltered; for RGB each row gets the
    /// filter with the smallest sum of absolute differences, the usual heuristic.
    /// </summary>
    private static byte[] Filter(int width, int height, byte[] packed, int bytesPerPixel, bool adaptive)
    {
        int stride = width * bytesPerPixel;
        byte[] result = new byte[(stride + 1) * height];
        byte[] candidate = new byte[stride];

        for (int row = 0; row < height; row++)
        {
            ReadOnlySpan<byte> line = packed.AsSpan(row * stride, stride);
            ReadOnlySpan<byte> previous = ReadOnlySpan<byte>.Empty;

            if (row > 0)
            {
                previous = packed.AsSpan((row - 1) * stride, stride);
            }

            int bestFilter = 0;
            long bestScore = long.MaxValue;
            int last = 0;

            if (adaptive == true)
            {
                last = 4;
            }

            for (int filter = 0; filter <= last; filter++)
            {
                long score = Apply(filter, line, previous, bytesPerPixel, candidate);

                if (score < bestScore)
                {
                    bestScore = score;
                    bestFilter = filter;
                }
            }

            Apply(bestFilter, line, previous, bytesPerPixel, candidate);
            result[row * (stride + 1)] = (byte)bestFilter;
            candidate.CopyTo(result.AsSpan((row * (stride + 1)) + 1));
        }

        return result;
    }

    private static long Apply(int filter, ReadOnlySpan<byte> line, ReadOnlySpan<byte> previous, int bytesPerPixel, Span<byte> output)
    {
        long score = 0;

        for (int index = 0; index < line.Length; index++)
        {
            int left = 0;
            int up = 0;
            int upLeft = 0;

            if (index >= bytesPerPixel)
            {
                left = line[index - bytesPerPixel];
            }

            if (previous.IsEmpty == false)
            {
                up = previous[index];

                if (index >= bytesPerPixel)
                {
                    upLeft = previous[index - bytesPerPixel];
                }
            }

            int predicted = filter switch
            {
                1 => left,
                2 => up,
                3 => (left + up) / 2,
                4 => Paeth(left, up, upLeft),
                _ => 0,
            };

            byte value = (byte)(line[index] - predicted);
            output[index] = value;
            score += Math.Abs((int)(sbyte)value);
        }

        return score;
    }

    /// <summary>The CRC-32 a PNG chunk ends with, carried on from <paramref name="crc"/>.</summary>
    private static uint Crc(uint crc, ReadOnlySpan<byte> data)
    {
        uint result = crc;

        foreach (byte value in data)
        {
            result ^= value;

            for (int bit = 0; bit < 8; bit++)
            {
                uint mask = 0u - (result & 1u);
                result = (result >> 1) ^ (0xEDB88320u & mask);
            }
        }

        return result;
    }

    private static void WriteChunk(Stream output, string type, byte[] data)
    {
        Span<byte> length = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(length, data.Length);
        output.Write(length);
        byte[] typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        output.Write(typeBytes);
        output.Write(data);
        uint crc = Crc(Crc(0xFFFFFFFFu, typeBytes), data) ^ 0xFFFFFFFFu;
        Span<byte> checksum = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(checksum, crc);
        output.Write(checksum);
    }
}
