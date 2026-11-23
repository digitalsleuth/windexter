using System.Globalization;
using System.IO;

namespace Windexter.Core
{
    public static partial class SearchIndex
    {
        private const string SecStoreFileName = "SecStore.db";
        private const string SecurityDescriptorTable = "SecurityDescriptor";
        private const string SdidColumn = "SDID";
        private const string GatherSecurityDescriptorColumn = "SecurityDescriptor";
        private const string GatherOwnerColumn = "Owner";
        private const string DomainSidPrefix = "S-1-5-21-";
        private const string AccessMaskFormat = "0x{0:x8}";

        private static List<List<object>> securityStore = [];
        private static List<List<object>> SecurityResults = [];
        private static List<List<object>> SecurityEntryResults = [];
        private static bool SecStoreAvailable = false;

        private static readonly Dictionary<string, (string Sid, string Name)> SidAliases = new()
        {
            ["AN"] = ("S-1-5-7", "Anonymous"),
            ["AU"] = ("S-1-5-11", "Authenticated Users"),
            ["BA"] = ("S-1-5-32-544", "Administrators"),
            ["BG"] = ("S-1-5-32-546", "Guests"),
            ["BO"] = ("S-1-5-32-551", "Backup Operators"),
            ["BU"] = ("S-1-5-32-545", "Users"),
            ["CG"] = ("S-1-3-1", "Creator Group"),
            ["CO"] = ("S-1-3-0", "Creator Owner"),
            ["IU"] = ("S-1-5-4", "Interactive"),
            ["LS"] = ("S-1-5-19", "Local Service"),
            ["NS"] = ("S-1-5-20", "Network Service"),
            ["NU"] = ("S-1-5-2", "Network"),
            ["OW"] = ("S-1-3-4", "Owner Rights"),
            ["PS"] = ("S-1-5-10", "Principal Self"),
            ["PU"] = ("S-1-5-32-547", "Power Users"),
            ["RC"] = ("S-1-5-12", "Restricted Code"),
            ["RD"] = ("S-1-5-32-555", "Remote Desktop Users"),
            ["RE"] = ("S-1-5-32-552", "Replicator"),
            ["SO"] = ("S-1-5-32-549", "Server Operators"),
            ["SU"] = ("S-1-5-6", "Service"),
            ["SY"] = ("S-1-5-18", "Local System"),
            ["WD"] = ("S-1-1-0", "Everyone"),
            ["AC"] = ("S-1-15-2-1", "All Application Packages"),
            ["AA"] = ("S-1-5-32-579", "Access Control Assistance Operators"),
            ["ED"] = ("S-1-5-9", "Enterprise Domain Controllers"),
            ["IS"] = ("S-1-5-17", "IIS User")
        };

        private static readonly Dictionary<string, string> KnownSidNames = SidAliases.Values.ToDictionary(entry => entry.Sid, entry => entry.Name);

        private static readonly Dictionary<uint, string> AccountRidNames = new()
        {
            [500] = "Administrator account",
            [501] = "Guest account",
            [503] = "Default account",
            [504] = "WDAG utility account",
            [512] = "Domain Admins",
            [513] = "None / Domain Users",
            [514] = "Domain Guests"
        };

        private static readonly Dictionary<string, uint> RightAliases = new()
        {
            ["GA"] = 0x10000000, ["GR"] = 0x80000000, ["GW"] = 0x40000000, ["GX"] = 0x20000000,
            ["RC"] = 0x00020000, ["SD"] = 0x00010000, ["WD"] = 0x00040000, ["WO"] = 0x00080000,
            ["RP"] = 0x00000010, ["WP"] = 0x00000020, ["CC"] = 0x00000001, ["DC"] = 0x00000002,
            ["LC"] = 0x00000004, ["SW"] = 0x00000008, ["LO"] = 0x00000080, ["DT"] = 0x00000040,
            ["CR"] = 0x00000100, ["FA"] = 0x001f01ff, ["FR"] = 0x00120089, ["FW"] = 0x00120116,
            ["FX"] = 0x001200a0, ["KA"] = 0x000f003f, ["KR"] = 0x00020019, ["KW"] = 0x00020006,
            ["KX"] = 0x00020019, ["NW"] = 0x00000001, ["NR"] = 0x00000002, ["NX"] = 0x00000004
        };

        private static readonly (uint Mask, string Name)[] KnownAccessMasks =
        [
            (0x001f01ff, "Full control"),
            (0x001301bf, "Modify"),
            (0x001200a9, "Read and execute"),
            (0x00120089, "Read"),
            (0x00100116, "Write"),
            (0x00120116, "Write"),
            (0x001200a0, "Execute")
        ];

        private static readonly (uint Mask, string Name)[] AccessBits =
        [
            (0x80000000, "GenericRead"), (0x40000000, "GenericWrite"), (0x20000000, "GenericExecute"), (0x10000000, "GenericAll"),
            (0x00100000, "Synchronize"), (0x00080000, "WriteOwner"), (0x00040000, "WriteDac"), (0x00020000, "ReadControl"), (0x00010000, "Delete"),
            (0x00000100, "WriteAttributes"), (0x00000080, "ReadAttributes"), (0x00000040, "DeleteChild"), (0x00000020, "Execute/Traverse"),
            (0x00000010, "WriteEA"), (0x00000008, "ReadEA"), (0x00000004, "AppendData/AddSubdirectory"), (0x00000002, "WriteData/AddFile"), (0x00000001, "ReadData/ListDirectory")
        ];

        private static readonly Dictionary<string, string> AceTypeNames = new()
        {
            ["A"] = "Allow", ["D"] = "Deny", ["OA"] = "Allow (object)", ["OD"] = "Deny (object)", ["AU"] = "Audit", ["AL"] = "Alarm",
            ["OU"] = "Audit (object)", ["OL"] = "Alarm (object)", ["ML"] = "Mandatory label", ["XA"] = "Allow (conditional)", ["XD"] = "Deny (conditional)",
            ["RA"] = "Resource attribute", ["SP"] = "Scoped policy", ["XU"] = "Audit (conditional)", ["TL"] = "Trust label", ["FL"] = "Access filter"
        };

        private static readonly Dictionary<string, string> AceFlagNames = new()
        {
            ["OI"] = "Object inherit", ["CI"] = "Container inherit", ["IO"] = "Inherit only", ["NP"] = "No propagate", ["ID"] = "Inherited", ["SA"] = "Audit success", ["FA"] = "Audit failure"
        };

        private static readonly Dictionary<string, string> ControlFlagNames = new()
        {
            ["P"] = "Protected", ["AR"] = "Auto-inherit required", ["AI"] = "Auto-inherited", ["NO_ACCESS_CONTROL"] = "No access control"
        };

        public static string? FindSecStoreFile() => dbType == "index" ? SearchBelow(dbFile, SecStoreFileName) : null;

        public static string? FindSecStoreFile(string path)
        {
            string name = Path.GetFileName(path);
            return !name.EndsWith(".db", StringComparison.OrdinalIgnoreCase) || name.Contains("gather", StringComparison.OrdinalIgnoreCase) || name.Equals("PropMap.db", StringComparison.OrdinalIgnoreCase) || name.Equals(SecStoreFileName, StringComparison.OrdinalIgnoreCase) ? null : SearchBelow(path, SecStoreFileName);
        }

        public static void ExtractSecStore(string secStoreFile)
        {
            ExtractCompanion(secStoreFile);
            SecStoreAvailable = true;
        }

        private static string DescribeSid(string value)
        {
            string sid = SidAliases.TryGetValue(value, out var alias) ? alias.Sid : value;
            if (KnownSidNames.TryGetValue(sid, out string? name))
            {
                return name;
            }
            string[] parts = sid.Split('-');
            if (sid.StartsWith(DomainSidPrefix, StringComparison.Ordinal) && parts.Length == 8 && uint.TryParse(parts[^1], NumberStyles.None, CultureInfo.InvariantCulture, out uint rid))
            {
                return AccountRidNames.TryGetValue(rid, out string? account) ? account : $"Account with RID {rid}";
            }
            if (sid.StartsWith("S-1-15-3-", StringComparison.Ordinal))
            {
                return "App capability";
            }
            if (sid.StartsWith("S-1-15-2-", StringComparison.Ordinal))
            {
                return "App package";
            }
            if (sid.StartsWith("S-1-5-80-", StringComparison.Ordinal))
            {
                return "Service";
            }
            return sid.StartsWith("S-1-5-5-", StringComparison.Ordinal) ? "Logon session" : "";
        }

        private static string ExpandSid(string value) => SidAliases.TryGetValue(value, out var alias) ? alias.Sid : value;

        private static string TrusteeText(string value)
        {
            string name = DescribeSid(value);
            return name.Length == 0 ? ExpandSid(value) : $"{name} ({ExpandSid(value)})";
        }

        private static (uint Mask, string Text) ParseRights(string text)
        {
            uint mask = 0;
            if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase) && uint.TryParse(text.AsSpan(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint hex))
            {
                mask = hex;
            }
            else if (uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out uint number))
            {
                mask = number;
            }
            else
            {
                for (int index = 0; index + 2 <= text.Length; index += 2)
                {
                    if (RightAliases.TryGetValue(text.Substring(index, 2), out uint bit))
                    {
                        mask |= bit;
                    }
                }
            }
            return (mask, DescribeMask(mask));
        }

        private static string DescribeMask(uint mask)
        {
            foreach (var (known, name) in KnownAccessMasks)
            {
                if (mask == known)
                {
                    return name;
                }
            }
            string bits = string.Join(", ", AccessBits.Where(bit => (mask & bit.Mask) != 0).Select(bit => bit.Name));
            return bits.Length == 0 ? "None" : $"Special: {bits}";
        }

        private static string DescribeControlFlags(string flags)
        {
            List<string> parts = [];
            int index = 0;
            while (index < flags.Length)
            {
                string code = flags[index..].StartsWith("NO_ACCESS_CONTROL", StringComparison.Ordinal) ? "NO_ACCESS_CONTROL" : flags[index..].StartsWith("AR", StringComparison.Ordinal) || flags[index..].StartsWith("AI", StringComparison.Ordinal) ? flags.Substring(index, 2) : flags.Substring(index, 1);
                parts.Add(ControlFlagNames.GetValueOrDefault(code, code));
                index += code.Length;
            }
            return string.Join(", ", parts);
        }

        private static string DescribeFlags(string flags, Dictionary<string, string> names)
        {
            List<string> parts = [];
            for (int index = 0; index + 2 <= flags.Length; index += 2)
            {
                string code = flags.Substring(index, 2);
                parts.Add(names.TryGetValue(code, out string? name) ? name : code);
            }
            return string.Join(", ", parts);
        }

        private sealed record ParsedAce(string Section, string Type, string Flags, string Rights, string ObjectType, string InheritedObjectType, string Trustee, string Condition);

        private sealed record ParsedDescriptor(string Owner, string Group, string DaclFlags, string SaclFlags, List<ParsedAce> Entries);

        private static string ReadSddlSid(string text, ref int position)
        {
            int start = position;
            if (position + 1 < text.Length && text[position] == 'S' && text[position + 1] == '-')
            {
                while (position < text.Length && (char.IsDigit(text[position]) || text[position] == 'S' || text[position] == '-'))
                {
                    position++;
                }
                return text[start..position];
            }
            position = Math.Min(text.Length, position + 2);
            return text[start..position];
        }

        private static ParsedDescriptor? ParseSddl(string text)
        {
            string owner = "", group = "", daclFlags = "", saclFlags = "";
            List<ParsedAce> entries = [];
            int position = 0;
            while (position + 1 < text.Length)
            {
                char section = text[position];
                if (text[position + 1] != ':' || section is not ('O' or 'G' or 'D' or 'S'))
                {
                    return null;
                }
                position += 2;
                if (section is 'O' or 'G')
                {
                    string sid = ReadSddlSid(text, ref position);
                    if (section == 'O')
                    {
                        owner = sid;
                    }
                    else
                    {
                        group = sid;
                    }
                    continue;
                }
                int flagStart = position;
                while (position < text.Length && text[position] != '(' && !(position + 1 < text.Length && text[position + 1] == ':' && text[position] is 'O' or 'G' or 'D' or 'S'))
                {
                    position++;
                }
                string flags = text[flagStart..position];
                if (section == 'D')
                {
                    daclFlags = flags;
                }
                else
                {
                    saclFlags = flags;
                }
                while (position < text.Length && text[position] == '(')
                {
                    int depth = 0, aceStart = position + 1;
                    do
                    {
                        depth += text[position] == '(' ? 1 : text[position] == ')' ? -1 : 0;
                        position++;
                    } while (position < text.Length && depth > 0);
                    if (depth != 0)
                    {
                        return null;
                    }
                    string[] fields = text[aceStart..(position - 1)].Split(';');
                    if (fields.Length < 6)
                    {
                        return null;
                    }
                    entries.Add(new ParsedAce(section == 'D' ? "DACL" : "SACL", fields[0], fields[1], fields[2], fields[3], fields[4], fields[5], string.Join(";", fields.Skip(6))));
                }
            }
            return position == text.Length ? new ParsedDescriptor(owner, group, daclFlags, saclFlags, entries) : null;
        }

        private static void ParseSecurityDescriptors()
        {
            SecurityResults =
            [
                ["Id", "Owner", "Group", "DaclFlags", "SaclFlags", "AccessEntries", "SDDL", "EnterpriseIds"]
            ];
            SecurityEntryResults =
            [
                ["SecurityDescriptorId", "Entry", "Section", "Type", "Trustee", "TrusteeName", "Access", "AccessMask", "Flags", "Inheritance", "ObjectType", "InheritedObjectType", "Condition"]
            ];
            int idIndex = securityStore.Count > 1 ? securityStore[0].IndexOf("Id") : -1;
            int valueIndex = securityStore.Count > 1 ? securityStore[0].IndexOf("Value") : -1;
            int enterpriseIndex = securityStore.Count > 1 ? securityStore[0].IndexOf("EnterpriseIds") : -1;
            if (idIndex < 0 || valueIndex < 0)
            {
                return;
            }
            foreach (var row in securityStore.Skip(1).OrderBy(row => ToKey(row[idIndex]) ?? long.MaxValue))
            {
                object id = row[idIndex];
                string sddl = (row[valueIndex]?.ToString() ?? "").TrimEnd('\0');
                object enterprise = enterpriseIndex >= 0 && row[enterpriseIndex] is not (null or DBNull) ? row[enterpriseIndex] : DBNull.Value;
                ParsedDescriptor? parsed = ParseSddl(sddl);
                if (parsed is null)
                {
                    SecurityResults.Add([id, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, sddl, enterprise]);
                    continue;
                }
                SecurityResults.Add([id, TrusteeText(parsed.Owner), TrusteeText(parsed.Group), DescribeControlFlags(parsed.DaclFlags), DescribeControlFlags(parsed.SaclFlags), parsed.Entries.Count, sddl, enterprise]);
                int entry = 0;
                foreach (ParsedAce ace in parsed.Entries)
                {
                    (uint mask, string access) = ParseRights(ace.Rights);
                    SecurityEntryResults.Add([id, ++entry, ace.Section, AceTypeNames.GetValueOrDefault(ace.Type, ace.Type), ExpandSid(ace.Trustee), DescribeSid(ace.Trustee), access, string.Format(CultureInfo.InvariantCulture, AccessMaskFormat, mask), ace.Flags, DescribeFlags(ace.Flags, AceFlagNames), ace.ObjectType, ace.InheritedObjectType, ace.Condition]);
                }
            }
        }

        private static void AddGatherSecurity()
        {
            int sdidIndex = GatherResults.Count > 1 ? GatherResults[0].IndexOf(SdidColumn) : -1;
            if (sdidIndex < 0 || SecurityResults.Count < 2)
            {
                return;
            }
            Dictionary<long, (string Sddl, object Owner)> lookup = [];
            foreach (var row in SecurityResults.Skip(1))
            {
                if (ToKey(row[0]) is long id)
                {
                    lookup.TryAdd(id, (row[6]?.ToString() ?? "", row[1]));
                }
            }
            GatherResults[0].InsertRange(sdidIndex + 1, [GatherSecurityDescriptorColumn, GatherOwnerColumn]);
            foreach (var row in GatherResults.Skip(1))
            {
                bool found = row.Count > sdidIndex && ToKey(row[sdidIndex]) is long id && lookup.ContainsKey(id);
                (string Sddl, object Owner) match = found ? lookup[ToKey(row[sdidIndex])!.Value] : default;
                row.InsertRange(Math.Min(sdidIndex + 1, row.Count), [found ? match.Sddl : DBNull.Value, found ? match.Owner : DBNull.Value]);
            }
        }
    }
}
