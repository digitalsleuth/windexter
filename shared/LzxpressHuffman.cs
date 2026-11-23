namespace Windexter.Core
{
    internal static class LzxpressHuffman
    {
        private const int SymbolCount = 512;
        private const int LiteralSymbols = 256;
        private const int MaxCodeLength = 15;
        private const int MaxOutputSize = 65535;
        private const int MinimumMatch = 3;
        private const int LengthMask = 15;
        private const int OffsetShift = 4;
        private const int LongLength = 15;
        private const int LongLengthEscape = 255;
        private const int WordBits = 16;

        public static byte[]? Decompress(byte[] input, int start, int outputSize)
        {
            if (outputSize <= 0 || outputSize > MaxOutputSize || input.Length - start < SymbolCount / 2)
            {
                return null;
            }
            int[] counts = new int[MaxCodeLength + 1];
            int[] lengths = new int[SymbolCount];
            for (int symbol = 0; symbol < SymbolCount; symbol++)
            {
                byte packed = input[start + symbol / 2];
                lengths[symbol] = symbol % 2 == 0 ? packed & 15 : packed >> 4;
                counts[lengths[symbol]]++;
            }
            counts[0] = 0;
            int[] first = new int[MaxCodeLength + 2];
            int[] offsets = new int[MaxCodeLength + 2];
            int remaining = 1;
            for (int length = 1; length <= MaxCodeLength; length++)
            {
                remaining = remaining * 2 - counts[length];
                if (remaining < 0)
                {
                    return null;
                }
                first[length + 1] = (first[length] + counts[length]) << 1;
                offsets[length + 1] = offsets[length] + counts[length];
            }
            int[] sorted = new int[SymbolCount];
            int[] next = (int[])offsets.Clone();
            for (int symbol = 0; symbol < SymbolCount; symbol++)
            {
                if (lengths[symbol] != 0)
                {
                    sorted[next[lengths[symbol]]++] = symbol;
                }
            }
            Bits bits = new(input, start + SymbolCount / 2);
            byte[] output = new byte[outputSize];
            int written = 0;
            while (written < outputSize)
            {
                int peeked = bits.Peek(MaxCodeLength);
                int symbol = -1;
                for (int length = 1; length <= MaxCodeLength; length++)
                {
                    int code = peeked >> (MaxCodeLength - length);
                    if (code >= first[length] && code - first[length] < counts[length])
                    {
                        symbol = sorted[offsets[length] + code - first[length]];
                        bits.Skip(length);
                        break;
                    }
                }
                if (symbol < 0)
                {
                    return null;
                }
                if (symbol < LiteralSymbols)
                {
                    output[written++] = (byte)symbol;
                    continue;
                }
                symbol -= LiteralSymbols;
                int matchLength = symbol & LengthMask;
                int offsetBits = symbol >> OffsetShift;
                if (matchLength == LongLength)
                {
                    int extra = bits.ReadByte();
                    if (extra == LongLengthEscape)
                    {
                        extra = bits.ReadUInt16();
                        matchLength = extra == 0 ? -1 : extra;
                    }
                    else
                    {
                        matchLength = extra < 0 ? -1 : extra + LongLength;
                    }
                    if (matchLength < 0)
                    {
                        return null;
                    }
                }
                int offset = offsetBits == 0 ? 1 : 1 << offsetBits | bits.Peek(offsetBits);
                bits.Skip(offsetBits);
                matchLength += MinimumMatch;
                if (offset > written || matchLength > outputSize - written || bits.Overrun)
                {
                    return null;
                }
                for (int copied = 0; copied < matchLength; copied++, written++)
                {
                    output[written] = output[written - offset];
                }
            }
            return output;
        }

        private sealed class Bits(byte[] input, int position)
        {
            private uint buffer = (uint)ReadWord(input, ref position) << WordBits | ReadWord(input, ref position);
            private int count = 2 * WordBits;
            private int overrunWords;

            public bool Overrun => overrunWords > 2 || position > input.Length + 4;

            public int Peek(int bits) => bits == 0 ? 0 : (int)(buffer >> (2 * WordBits - bits));

            public void Skip(int bits)
            {
                buffer <<= bits;
                count -= bits;
                if (count < WordBits)
                {
                    if (position + 2 > input.Length)
                    {
                        overrunWords++;
                    }
                    buffer |= (uint)ReadWord(input, ref position) << (WordBits - count);
                    count += WordBits;
                }
            }

            public int ReadByte()
            {
                if (position >= input.Length)
                {
                    return -1;
                }
                return input[position++];
            }

            public int ReadUInt16()
            {
                if (position + 2 > input.Length)
                {
                    position = input.Length + 5;
                    return 0;
                }
                return ReadWord(input, ref position);
            }

            private static ushort ReadWord(byte[] data, ref int index)
            {
                ushort word = index + 1 < data.Length ? (ushort)(data[index] | data[index + 1] << 8) : (ushort)0;
                index += 2;
                return word;
            }
        }
    }
}
