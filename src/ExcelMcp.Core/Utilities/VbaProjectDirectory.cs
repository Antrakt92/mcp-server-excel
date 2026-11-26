using System.Buffers.Binary;
using System.Numerics;
using System.Security.Cryptography;
using System.Text;

namespace Sbroenne.ExcelMcp.Core.Utilities;

// The project directory grammar and compression are described in MS-OVBA sections 2.3.4.2 and 2.4.1.
internal static class VbaProjectDirectory
{
    private static readonly UnicodeEncoding Utf16 = new(bigEndian: false, byteOrderMark: false, throwOnInvalidBytes: true);

    internal sealed record Module(string Name, string StreamName, int SourceOffset, string Metadata);
    internal sealed record Project(Encoding Encoding, string Metadata, string References, int ReferenceCount, IReadOnlyList<Module> Modules);

    internal static Project Read(byte[] compressed, CancellationToken cancellationToken)
    {
        var reader = new RecordReader(Decompress(compressed, 1024 * 1024, cancellationToken));
        reader.Record(0x0001, 4); // System kind
        if (reader.PeekId() == 0x004A)
            reader.Record(0x004A, 4); // Optional compatibility version
        reader.Record(0x0002, 4);
        reader.Record(0x0014, 4);
        int codePage = BinaryPrimitives.ReadUInt16LittleEndian(reader.Record(0x0003, 2));
        Encoding encoding = codePage == 65001 ? new UTF8Encoding(false, true) :
            CodePagesEncodingProvider.Instance.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)
            ?? throw new InvalidDataException("The VBA project code page is unsupported.");
        reader.Record(0x0004);
        reader.Pair(0x0005, 0x0040);
        reader.Pair(0x0006, 0x003D);
        reader.Record(0x0007, 4);
        reader.Record(0x0008, 4);
        reader.Id(0x0009);
        reader.Length(4); // Reserved value; the version itself occupies six bytes.
        reader.Take(6);
        if (reader.PeekId() == 0x000C)
            reader.Pair(0x000C, 0x003C); // Optional project constants

        string metadata = Hash(reader.Slice(0, reader.Position));
        int referenceStart = reader.Position;
        int referenceCount = 0;
        while (reader.PeekId() != 0x000F)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (++referenceCount > 1024)
                throw new InvalidDataException("The VBA project reference limit was exceeded.");
            ReadReference(ref reader);
        }
        string references = Hash(reader.Slice(referenceStart, reader.Position));

        int moduleCount = BinaryPrimitives.ReadUInt16LittleEndian(reader.Record(0x000F, 2));
        if (moduleCount > 1024)
            throw new InvalidDataException("The VBA project module limit was exceeded.");
        reader.Record(0x0013, 2); // Ignored project cookie
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var streams = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var modules = new List<Module>(moduleCount);
        for (int index = 0; index < moduleCount; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var module = ReadModule(ref reader, encoding);
            if (!names.Add(module.Name) || !streams.Add(module.StreamName))
                throw new InvalidDataException("The VBA directory contains ambiguous module identities.");
            modules.Add(module);
        }
        reader.Record(0x0010, 0);
        if (!reader.End)
            throw new InvalidDataException("The VBA directory has unsupported trailing data.");
        return new(encoding, metadata, references, referenceCount, modules);
    }

    private static void ReadReference(ref RecordReader reader)
    {
        if (reader.PeekId() == 0x0016)
            reader.Pair(0x0016, 0x003E);
        if (reader.PeekId() == 0x0033)
        {
            reader.Record(0x0033);
            if (reader.PeekId() != 0x002F)
                throw new InvalidDataException("The VBA reference record is incomplete.");
        }
        switch (reader.PeekId())
        {
            case 0x002F: // ActiveX control (optional extended reference name)
                reader.Id(0x002F);
                reader.Take(4); // Ignore outer size fields as specified in MS-OVBA.
                reader.Take(reader.Length());
                reader.Take(6);
                if (reader.PeekId() == 0x0016)
                    reader.Pair(0x0016, 0x003E);
                reader.Id(0x0030);
                reader.Take(4);
                reader.Take(reader.Length());
                reader.Take(26);
                break;
            case 0x000D: // Registered type library
                reader.Id(0x000D);
                reader.Take(4);
                reader.Take(reader.Length());
                reader.Take(6);
                break;
            case 0x000E: // VBA project
                reader.Id(0x000E);
                reader.Take(4);
                reader.Take(reader.Length());
                reader.Take(reader.Length());
                reader.Take(6);
                break;
            default:
                throw new InvalidDataException("The VBA directory contains an unsupported reference.");
        }
    }

    private static Module ReadModule(ref RecordReader reader, Encoding encoding)
    {
        int start = reader.Position;
        string name = encoding.GetString(reader.Record(0x0019));
        if (reader.PeekId() == 0x0047)
            name = MatchUnicode(name, reader.Record(0x0047));
        string stream = encoding.GetString(reader.Record(0x001A));
        stream = MatchUnicode(stream, reader.Record(0x0032));
        reader.Pair(0x001C, 0x0048);
        int offsetRecord = reader.Position;
        uint offset = BinaryPrimitives.ReadUInt32LittleEndian(reader.Record(0x0031, 4));
        int afterOffset = reader.Position;
        reader.Record(0x001E, 4);
        int cookieRecord = reader.Position;
        reader.Record(0x002C, 2);
        int afterCookie = reader.Position;
        ushort kind = reader.PeekId();
        if (kind is not (0x0021 or 0x0022))
            throw new InvalidDataException("The VBA directory contains an unsupported module type.");
        reader.Record(kind, 0);
        if (reader.PeekId() == 0x0025)
            reader.Record(0x0025, 0);
        if (reader.PeekId() == 0x0028)
            reader.Record(0x0028, 0);
        reader.Record(0x002B, 0);
        if (!IsSafeName(name) || !IsSafeName(stream) || offset > int.MaxValue ||
            stream.Equals("dir", StringComparison.OrdinalIgnoreCase) ||
            stream.Equals("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The VBA directory contains an unsupported module identity.");
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        hash.AppendData(reader.Slice(start, offsetRecord));
        hash.AppendData(reader.Slice(afterOffset, cookieRecord));
        hash.AppendData(reader.Slice(afterCookie, reader.Position));
        return new(name, stream, (int)offset, Convert.ToHexString(hash.GetHashAndReset()));
    }

    private static string MatchUnicode(string name, ReadOnlySpan<byte> bytes)
    {
        if (bytes.IsEmpty)
            return name;
        string unicode = Utf16.GetString(bytes);
        return string.Equals(name, unicode, StringComparison.Ordinal) ? unicode :
            throw new InvalidDataException("The VBA directory contains inconsistent module names.");
    }

    internal static bool IsSafeName(string name) => name.Length is > 0 and <= 31 &&
        name.All(character => !char.IsControl(character) && character is not ('/' or '\\'));

    internal static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes));

    internal static byte[] Decompress(ReadOnlySpan<byte> bytes, int maxBytes, CancellationToken cancellationToken)
    {
        const int chunkSize = 4096;
        if (bytes.IsEmpty || bytes[0] != 1)
            throw new InvalidDataException("The VBA compressed container is invalid.");
        int position = 1;
        using var output = new MemoryStream();
        Span<byte> chunk = stackalloc byte[chunkSize];
        while (position < bytes.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (bytes.Length - position < 2)
                throw new InvalidDataException("The VBA compressed chunk header is incomplete.");
            ushort header = BinaryPrimitives.ReadUInt16LittleEndian(bytes[position..]);
            int length = (header & 0x0FFF) + 1; // Number of following payload bytes.
            position += 2;
            if ((header & 0x7000) != 0x3000 || length > bytes.Length - position)
                throw new InvalidDataException("The VBA compressed chunk header is invalid.");
            var payload = bytes.Slice(position, length);
            position += length;
            var expanded = (header & 0x8000) == 0 ? payload : chunk[..DecodeChunk(payload, chunk)];
            if (((header & 0x8000) == 0 || position < bytes.Length) && expanded.Length != chunkSize)
                throw new InvalidDataException("The VBA compressed chunk size is invalid.");
            if (expanded.Length > maxBytes - output.Length)
                throw new InvalidDataException("The VBA decompression limit was exceeded.");
            output.Write(expanded);
        }
        return output.ToArray();
    }

    private static int DecodeChunk(ReadOnlySpan<byte> payload, Span<byte> output)
    {
        int input = 0, written = 0;
        while (input < payload.Length)
        {
            byte flags = payload[input++];
            // MS-OVBA ends the token sequence at the chunk boundary. Excel can
            // write a last flag byte with no remaining tokens.
            for (int bit = 0; bit < 8 && input < payload.Length; bit++)
            {
                if (written == output.Length)
                    throw new InvalidDataException("The VBA compressed chunk expands past its limit.");
                if ((flags & (1 << bit)) == 0)
                {
                    output[written++] = payload[input++];
                    continue;
                }
                if (written == 0 || payload.Length - input < 2)
                    throw new InvalidDataException("The VBA compressed copy token is invalid.");
                ushort token = BinaryPrimitives.ReadUInt16LittleEndian(payload[input..]);
                input += 2;
                int offsetBits = Math.Max(4, written == 1 ? 0 : BitOperations.Log2((uint)(written - 1)) + 1);
                int lengthMask = ushort.MaxValue >> offsetBits;
                int length = (token & lengthMask) + 3;
                int distance = (token >> (16 - offsetBits)) + 1;
                if (distance > written || length > output.Length - written)
                    throw new InvalidDataException("The VBA compressed copy token is outside its chunk.");
                for (int index = 0; index < length; index++, written++)
                    output[written] = output[written - distance];
            }
        }
        return written;
    }

    private ref struct RecordReader
    {
        private readonly ReadOnlySpan<byte> _bytes;

        public RecordReader(ReadOnlySpan<byte> bytes) => _bytes = bytes;

        public int Position { get; private set; }
        public bool End => Position == _bytes.Length;
        public ReadOnlySpan<byte> Slice(int start, int end) => _bytes[start..end];

        public ushort PeekId() => _bytes.Length - Position >= 2 ? BinaryPrimitives.ReadUInt16LittleEndian(_bytes[Position..]) :
            throw new InvalidDataException("The VBA directory ended inside a record.");

        public void Id(ushort expected)
        {
            if (PeekId() != expected)
                throw new InvalidDataException("The VBA directory has an unsupported record layout.");
            Position += 2;
        }

        public int Length(int? expected = null)
        {
            uint length = BinaryPrimitives.ReadUInt32LittleEndian(Take(4));
            if (length > int.MaxValue || (expected is { } value && length != value))
                throw new InvalidDataException("The VBA directory record length is invalid.");
            return (int)length;
        }

        public ReadOnlySpan<byte> Take(int length)
        {
            if (length > _bytes.Length - Position)
                throw new InvalidDataException("The VBA directory ended inside a record.");
            var record = _bytes.Slice(Position, length);
            Position += length;
            return record;
        }

        public ReadOnlySpan<byte> Record(ushort id, int? length = null)
        {
            Id(id);
            return Take(Length(length));
        }

        public void Pair(ushort first, ushort second)
        {
            Record(first);
            Record(second);
        }
    }
}
