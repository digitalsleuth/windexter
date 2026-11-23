using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;

namespace Windexter.Core
{
    public static partial class SearchIndex
    {
        public class SerializedPropertyStore
        {
            public string? FormatId { get; set; }
            public Dictionary<string, object> Properties { get; set; } = [];
        }

        private const ushort BlobVariantType = 0x41;
        private const int MinimumStorageSize = 28;
        private const int VectorVariantFlag = 0x1000;
        private const int BooleanVariantType = 0x0B;

        private static readonly HashSet<int> VectorElementTypes = [0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x0B, 0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x40, 0x48];

        public static class SPSParser
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-propstore/3453fb82-0e4f-4c2c-bc04-64b4bd2c51ec
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/f122b9d7-e5cf-4484-8466-83f6fd94b3cc
        /// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/5a2b34c4-d109-438e-9ec8-84816d8de40d
        /// CodePage Property Id? - https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/04727953-c174-4d01-80ae-b38b15d7068a
        {
            public static List<SerializedPropertyStore> Parse(byte[] blob)
            {
                var stores = new List<SerializedPropertyStore>();
                try
                {
                    ParseInto(blob, stores);
                }
                catch (Exception ex) when (IsDecodeError(ex))
                {
                    if (stores.Count > 0)
                    {
                        var errorStore = new SerializedPropertyStore { FormatId = "Parse Error" };
                        errorStore.Properties["Error"] = ex.Message;
                        stores.Add(errorStore);
                    }
                }
                if (stores.Count == 0)
                {
                    var rawStore = new SerializedPropertyStore { FormatId = "Raw Binary Data" };
                    rawStore.Properties["Data"] = Convert.ToHexString(blob);
                    stores.Add(rawStore);
                }
                return stores;
            }

            private static bool IsDecodeError(Exception ex) => ex is EndOfStreamException or ArgumentException or InvalidCastException or InvalidDataException or InvalidOperationException or NullReferenceException or IndexOutOfRangeException or OverflowException;

            private static void ParseInto(byte[] blob, List<SerializedPropertyStore> stores)
            {
                using var ms = new MemoryStream(blob);
                using var reader = new BinaryReader(ms);
                var firstBytes = reader.ReadInt32();
                if (reader.BaseStream.Position < reader.BaseStream.Length - 4)
                {
                    if (firstBytes == 1)
                    {
                        reader.ReadInt32();
                        List<string> dataStrings = [];
                        string data;
                        while (reader.BaseStream.Position < reader.BaseStream.Length)
                        {
                            data = ReadUnicodeString(reader, false);
                            dataStrings.Add(data);
                            if (reader.BaseStream.Position != reader.BaseStream.Length && reader.BaseStream.Length - reader.BaseStream.Position != 4)
                            {
                                reader.BaseStream.Position -= 2;
                                reader.ReadInt16();
                            }
                        }
                        var store = new SerializedPropertyStore { FormatId = "Partially Serialized Data" };
                        store.Properties["Data"] = string.Join(" ", dataStrings);
                        stores.Add(store);
                        return;
                    }
                    else if ((blob[0] == 19 && blob[1] == 1) || (blob[0] == 20 && blob[1] == 1) || (blob[0] == 23 && blob[1] == 1))
                    {
                        string data;
                        List<string> dataStrings = [];
                        data = Decompress7Bit(blob);
                        dataStrings.Add(data.Replace('\0', ' ').Trim());
                        var store = new SerializedPropertyStore { FormatId = "7-bit Compressed String" };
                        store.Properties["Data"] = string.Join(" ", dataStrings);
                        stores.Add(store);
                        return;
                    }
                    int storeSize = reader.ReadInt32();
                    long endPos = reader.BaseStream.Position - 4 + storeSize;
                    while (reader.BaseStream.Position < endPos)
                    {
                        int storageSize = reader.ReadInt32();
                        long storageEnd = Math.Min(reader.BaseStream.Position - 4 + storageSize, blob.Length);
                        if (storageSize < MinimumStorageSize)
                        {
                            break;
                        }
                        string version = Encoding.ASCII.GetString(reader.ReadBytes(4));
                        if (version != "1SPS")
                        {
                            if (firstBytes == 6029404)
                            {
                                var rawStore = new SerializedPropertyStore { FormatId = "Raw Data" };
                                rawStore.Properties["Data"] = Encoding.Unicode.GetString(blob).Trim('\0');
                                stores.Add(rawStore);
                                return;
                            }
                            break;
                        }
                        string formatId = new Guid(reader.ReadBytes(16)).ToString("B");
                        while (reader.BaseStream.Position + 4 <= storageEnd)
                        {
                            long valueStart = reader.BaseStream.Position;
                            uint valueSize = reader.ReadUInt32();
                            long valueEnd = valueStart + valueSize;
                            if (valueSize == 0 || valueEnd > storageEnd)
                            {
                                break;
                            }
                            var store = new SerializedPropertyStore { FormatId = formatId };
                            try
                            {
                                ReadStoreValue(reader, store, formatId);
                            }
                            catch (Exception ex) when (IsDecodeError(ex))
                            {
                                store.Properties["Value"] = $"[Unable to decode: {ex.Message}]";
                            }
                            stores.Add(store);
                            reader.BaseStream.Seek(valueEnd, SeekOrigin.Begin);
                        }
                        reader.BaseStream.Seek(storageEnd, SeekOrigin.Begin);
                    }
                }
            }

            private const string StringNameFormatId = "{d5cdd505-2e9c-101b-9397-08002b2cf9ae}";

            private static void ReadStoreValue(BinaryReader reader, SerializedPropertyStore store, string formatId)
            {
                string propertyName = "";
                if (formatId == StringNameFormatId)
                {
                    int nameSize = reader.ReadInt32();
                    _ = reader.ReadBytes(1);
                    store.Properties["Name"] = Encoding.Unicode.GetString(reader.ReadBytes(nameSize)).TrimEnd('\0');
                }
                else
                {
                    uint spId = reader.ReadUInt32();
                    reader.ReadBytes(1);
                    if (PropertyKeyNames.TryGetValue((formatId, (int)spId), out string? knownName))
                    {
                        propertyName = knownName;
                        store.Properties["Name"] = propertyName;
                    }
                    store.Properties["PropertyId"] = spId.ToString();
                }
                ushort vt = reader.ReadUInt16();
                reader.ReadUInt16();
                if (vt == BlobVariantType)
                {
                    reader.ReadUInt32();
                    reader.ReadUInt16();
                    vt = reader.ReadUInt16();
                    reader.ReadUInt16();
                }
                object? spValue = ReadPropertyValue(reader, vt);
                if (propertyName == "Activity_ContentVisualPropertiesHash")
                {
                    spValue = BitConverter.ToString(BitConverter.GetBytes((ulong)spValue!)).Replace("-", "");
                }
                if (propertyName == "SFGAOFlags")
                {
                    spValue = string.Join("|", GetMatchingFlags((long)(uint)spValue!, SFGAO));
                }
                store.Properties["Value"] = spValue!;
                store.Properties["VT_ID"] = VariantType.GetValueOrDefault(vt, $"0x{vt:X}");
            }

            private static readonly JsonSerializerOptions writeOptions = new()
            {
                WriteIndented = true,
            };

            public static string ParseToJson(byte[] blob)
            {
                var stores = Parse(blob);
                return JsonSerializer.Serialize(stores, writeOptions);
            }

            private static object? ReadPropertyValue(BinaryReader reader, int vt)
            {
                /// Not all VT's are set here because they are not all used in the Windows Search Index.
                return vt switch
                {
                    0x00 => null, // VT_EMPTY
                    0x01 => null, // VT_NULL
                    0x02 => reader.ReadInt16(), // VT_I2
                    0x03 => reader.ReadInt32(), // VT_I4
                    0x04 => reader.ReadSingle(), // VT_R4
                    0x05 => reader.ReadDouble(), // VT_R8
                    0x06 => reader.ReadInt64(), // VT_CY
                    0x07 => reader.ReadDouble(), // VT_DATE
                    0x08 => ReadString(reader), // VT_BSTR
                    0x0A => reader.ReadUInt32(), // VT_ERROR
                    0x0B => reader.ReadInt32() != 0, // VT_BOOL
                    0x0E => ReadDecimal(reader), // VT_DECIMAL
                    0x10 => reader.ReadSByte(), // VT_I1
                    0x11 => reader.ReadByte(), // VT_UI1
                    0x12 => reader.ReadUInt16(), // VT_UI2
                    0x13 => reader.ReadUInt32(), // VT_UI4
                    0x14 => reader.ReadInt64(), // VT_I8
                    0x15 => reader.ReadUInt64(), // VT_UI8
                    0x16 => reader.ReadInt32(), // VT_INT
                    0x17 => reader.ReadUInt32(), // VT_UINT
                    0x1E => ReadString(reader), // VT_LPSTR
                    0x1F => ReadUnicodeString(reader, true), // VT_LPWSTR
                    0x40 => ReadFileTime(reader), // VT_FILETIME
                    0x41 => ReadBlob(reader), // VT_BLOB
                    0x42 => ReadUnicodeString(reader, true), // VT_STREAM
                    0x43 => ReadUnicodeString(reader, true), // VT_STORAGE
                    0x44 => ReadUnicodeString(reader, true), // VT_STREAMED_Object
                    0x45 => ReadUnicodeString(reader, true), // VT_STORED_Object
                    0x46 => ReadBlob(reader), // VT_BLOB_Object
                    0x47 => ReadClipboardData(reader), // VT_CF
                    0x48 => ReadGuid(reader), // VT_CLSID
                    0x49 => ReadVersionedStream(reader), // VT_VERSIONED_STREAM
                    0x101F => ReadVectorLpwstr(reader), // VT_VECTOR | VT_LPWSTR
                    _ when (vt & VectorVariantFlag) != 0 && VectorElementTypes.Contains(vt & ~VectorVariantFlag) => ReadVector(reader, vt & ~VectorVariantFlag),
                    _ => $"[Unsupported VT {vt:X}]"
                };
            }

            private static string ReadVector(BinaryReader reader, int elementType)
            {
                uint count = reader.ReadUInt32();
                if (count > reader.BaseStream.Length - reader.BaseStream.Position)
                {
                    throw new InvalidDataException($"The vector claims {count} elements but the data is too short.");
                }
                List<string> items = [];
                for (uint index = 0; index < count; index++)
                {
                    object? item = elementType == BooleanVariantType ? reader.ReadInt16() != 0 : ReadPropertyValue(reader, elementType);
                    items.Add(Convert.ToString(item, CultureInfo.InvariantCulture) ?? string.Empty);
                }
                return string.Join(" ", items);
            }

            private static DateTime ReadFileTime(BinaryReader reader)
            {
                return DateTime.FromFileTimeUtc(reader.ReadInt64());
            }

            private static byte[] ReadChecked(BinaryReader reader, long length)
            {
                if (length < 0 || length > reader.BaseStream.Length - reader.BaseStream.Position)
                {
                    throw new InvalidDataException($"A length of {length} runs past the end of the data.");
                }
                return reader.ReadBytes((int)length);
            }

            private static Guid ReadGuid(BinaryReader reader)
            {
                return new Guid(ReadChecked(reader, 16));
            }

            private static string ReadVectorLpwstr(BinaryReader reader)
            {
                // 2-byte buffers of null at the end of each property
                // Some entries don't align properly to the 2-byte buffer at the end, 
                // so an assessment has to be made if the next bytes are 0's and
                // re-align if necessary.
                var result = new List<string>();
                uint count = reader.ReadUInt32();
                long streamLength = reader.BaseStream.Length;
                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        uint charCount = reader.ReadUInt32();
                        byte[] strBytes = ReadChecked(reader, (long)charCount * 2);
                        string value = Encoding.Unicode.GetString(strBytes).TrimEnd('\0');
                        result.Add(value);
                        long pos = reader.BaseStream.Position;
                        if (pos != streamLength)
                        {
                            long buffer = BitConverter.ToInt16(reader.ReadBytes(2));
                            if (buffer != 0)
                            {
                                reader.BaseStream.Seek(pos, SeekOrigin.Begin);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"An error has occurred while attempting to read a VectorLpwstr:\n\n{ex.Message}");
                    }
                }
                string stringVal = string.Join(" ", result);
                return stringVal;
            }

            private static string ReadVersionedStream(BinaryReader reader)
            {
                var guid = new Guid(ReadChecked(reader, 16));
                uint charCount = reader.ReadUInt32();
                byte[] strBytes = ReadChecked(reader, (long)charCount * 2);
                string streamName = Encoding.Unicode.GetString(strBytes);
                long pos = reader.BaseStream.Position;
                long aligned = (pos + 3) & ~3;
                reader.BaseStream.Seek(aligned, SeekOrigin.Begin);
                return $"{guid} - {streamName}";

            }

            private static string ReadClipboardData(BinaryReader reader)
            {
                int length = reader.ReadInt32();
                var bytes = ReadChecked(reader, length);
                return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            }

            private static string ReadUnicodeString(BinaryReader reader, bool dbl)
            {
                int length = reader.ReadInt32();
                var bytes = ReadChecked(reader, dbl ? (long)length * 2 : length);
                return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            }

            private static string ReadString(BinaryReader reader)
            {
                int length = reader.ReadInt32();
                var bytes = ReadChecked(reader, length);
                return Encoding.ASCII.GetString(bytes).TrimEnd('\0');
            }

            private static string ReadBlob(BinaryReader reader)
            {
                int length = reader.ReadInt32();
                var bytes = ReadChecked(reader, length);
                return BitConverter.ToString(bytes).Replace("-", "").TrimEnd('\0');
            }

            public static decimal ReadDecimal(BinaryReader reader)
            {
                byte[] bytes = reader.ReadBytes(16);
                if (bytes.Length < 16)
                    throw new ArgumentException($"Need at least 16 bytes, only {bytes.Length} available.");
                ushort _ = BitConverter.ToUInt16(bytes, 0);
                byte scale = bytes[2];
                byte sign = bytes[3];
                uint hi32 = BitConverter.ToUInt32(bytes, 4);
                ulong lo64 = BitConverter.ToUInt64(bytes, 8);

                int lo = (int)(lo64 & 0xFFFFFFFF);
                int mid = (int)((lo64 >> 32) & 0xFFFFFFFF);
                int hi = (int)hi32;

                bool isNegative = (sign & 0x80) != 0;

                return new decimal(lo, mid, hi, isNegative, scale);
            }
        }
    }
}
