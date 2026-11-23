namespace Windexter.Core
{
    public static partial class SearchIndex
    {
        private static readonly string[] WorkIdColumns = ["WorkId", "DocID"];
        private const string DocumentIdColumn = "DocumentID";
        private const string GatherModifiedColumn = "LastModified";
        private const string InGatherColumn = "InGather";
        private const string GatherLastModifiedColumn = "GatherLastModified";
        private const string InIndexColumn = "InIndex";
        private const string PropertyMapIdColumn = "Id";
        private const string PropertyMapFormatColumn = "FormatId";
        private const string PropertyMapQualifierColumn = "FormatIdQualifier";
        private const string PropertyNameColumn = "PropertyName";
        private const string InPropertyStoreColumn = "InPropertyStore";
        private const string SystemPropertyPrefix = "System.";
        private const string PropertyMapStandardIdColumn = "StandardId";
        private const string InvertedOnlyPidsColumn = "InvertedOnlyPids";
        private const string InvertedOnlyPidNamesColumn = "InvertedOnlyPidNames";
        private const int MetadataNameIndex = 2;
        private const int MetadataPropertyIdIndex = 3;

        private static long? ToKey(object? value) => value switch
        {
            long number => number,
            int number => number,
            uint number => number,
            short number => number,
            ulong number when number <= long.MaxValue => (long)number,
            _ => null
        };

        private static string? ResolvePropertyName(string guid, long id)
        {
            if (id < 0 || id > int.MaxValue)
            {
                return null;
            }
            if (PropertyKeyNames.TryGetValue((guid, (int)id), out string? known))
            {
                return SystemPropertyPrefix + known.Replace('_', '.');
            }
            if (LegacyPropertyNames.TryGetValue((guid, (int)id), out string? legacy))
            {
                return legacy;
            }
            if (MapiNamedProperties.TryGetValue((guid, (int)id), out string? named))
            {
                return named;
            }
            return guid == MapiPropertySet && MapiTagProperties.TryGetValue((int)id, out string? tag) ? tag : null;
        }

        private static string? LabelPropertySet(string guid, long id) => PropertySetNames.TryGetValue(guid, out string? setName) ? $"{setName} {id}" : null;

        private static void AddGatherCorrelation()
        {
            int workIndex = IndexResults.Count > 1 ? IndexResults[0].FindIndex(header => header is string name && WorkIdColumns.Contains(name, StringComparer.OrdinalIgnoreCase)) : -1;
            int documentIndex = GatherResults.Count > 1 ? GatherResults[0].IndexOf(DocumentIdColumn) : -1;
            if (workIndex < 0 || documentIndex < 0)
            {
                return;
            }
            int modifiedIndex = GatherResults[0].IndexOf(GatherModifiedColumn);
            Dictionary<long, object> gathered = [];
            foreach (var row in GatherResults.Skip(1))
            {
                if (row.Count > documentIndex && ToKey(row[documentIndex]) is long document)
                {
                    gathered.TryAdd(document, modifiedIndex >= 0 && row.Count > modifiedIndex ? row[modifiedIndex] ?? DBNull.Value : DBNull.Value);
                }
            }
            HashSet<long> indexed = [.. IndexResults.Skip(1).Select(row => row.Count > workIndex ? ToKey(row[workIndex]) : null).OfType<long>()];
            if (!indexed.Overlaps(gathered.Keys))
            {
                return;
            }
            IndexResults[0].InsertRange(workIndex + 1, [InGatherColumn, GatherLastModifiedColumn]);
            foreach (var row in IndexResults.Skip(1))
            {
                bool found = row.Count > workIndex && ToKey(row[workIndex]) is long work && gathered.ContainsKey(work);
                row.InsertRange(Math.Min(workIndex + 1, row.Count), [found, found ? gathered[ToKey(row[workIndex])!.Value] : DBNull.Value]);
            }
            GatherResults[0].Insert(documentIndex + 1, InIndexColumn);
            foreach (var row in GatherResults.Skip(1))
            {
                row.Insert(Math.Min(documentIndex + 1, row.Count), row.Count > documentIndex && ToKey(row[documentIndex]) is long document && indexed.Contains(document));
            }
        }

        private static void AddPropertyMapNames()
        {
            int idIndex = PropertyMapResults.Count > 1 ? PropertyMapResults[0].IndexOf(PropertyMapIdColumn) : -1;
            int formatIndex = idIndex >= 0 ? PropertyMapResults[0].IndexOf(PropertyMapFormatColumn) : -1;
            int qualifierIndex = idIndex >= 0 ? PropertyMapResults[0].IndexOf(PropertyMapQualifierColumn) : -1;
            if (formatIndex < 0 || qualifierIndex < 0)
            {
                return;
            }
            bool hasStore = propertyMetadata.Count > 1;
            Dictionary<long, string> stored = hasStore ? propertyMetadata.Skip(1).Where(row => row.Count > MetadataPropertyIdIndex && ToKey(row[MetadataPropertyIdIndex]) is not null && row[MetadataNameIndex] is string).GroupBy(row => ToKey(row[MetadataPropertyIdIndex])!.Value).ToDictionary(group => group.Key, group => (string)group.First()[MetadataNameIndex]) : [];
            PropertyMapResults[0].InsertRange(idIndex + 1, hasStore ? [PropertyNameColumn, InPropertyStoreColumn] : [PropertyNameColumn]);
            foreach (var row in PropertyMapResults.Skip(1))
            {
                if (row.Count <= Math.Max(idIndex, Math.Max(formatIndex, qualifierIndex)))
                {
                    continue;
                }
                bool isStored = ToKey(row[idIndex]) is long id && stored.ContainsKey(id);
                object name = isStored ? stored[ToKey(row[idIndex])!.Value] : row[qualifierIndex] switch
                {
                    string text => text,
                    object number when ToKey(number) is long pid && row[formatIndex] is string format && (ResolvePropertyName(format.ToLowerInvariant(), pid) ?? LabelPropertySet(format.ToLowerInvariant(), pid)) is string resolved => resolved,
                    _ => DBNull.Value
                };
                row.InsertRange(idIndex + 1, hasStore ? [name, isStored] : [name]);
            }
        }

        private static void AddInvertedOnlyNames()
        {
            int pidIndex = IndexResults.Count > 1 ? IndexResults[0].IndexOf(InvertedOnlyPidsColumn) : -1;
            int idIndex = PropertyMapResults.Count > 1 ? PropertyMapResults[0].IndexOf(PropertyMapIdColumn) : -1;
            int nameIndex = idIndex >= 0 ? PropertyMapResults[0].IndexOf(PropertyNameColumn) : -1;
            if (pidIndex < 0 || nameIndex < 0)
            {
                return;
            }
            int standardIndex = PropertyMapResults[0].IndexOf(PropertyMapStandardIdColumn);
            Dictionary<long, string> names = [];
            foreach (var row in PropertyMapResults.Skip(1).Where(row => row.Count > nameIndex && row[nameIndex] is string))
            {
                string name = (string)row[nameIndex];
                if (ToKey(row[idIndex]) is long id)
                {
                    names.TryAdd(id, name);
                }
                if (standardIndex >= 0 && row.Count > standardIndex && ToKey(row[standardIndex]) is long standard)
                {
                    names.TryAdd(standard, name);
                }
            }
            IndexResults[0].Insert(pidIndex + 1, InvertedOnlyPidNamesColumn);
            foreach (var row in IndexResults.Skip(1))
            {
                row.Insert(Math.Min(pidIndex + 1, row.Count), row.Count > pidIndex && row[pidIndex] is string { Length: > 0 } pids ? string.Join("|", pids.Split('|').Select(pid => long.TryParse(pid, out long key) && names.TryGetValue(key, out string? known) ? known : pid)) : DBNull.Value);
            }
        }
    }
}
