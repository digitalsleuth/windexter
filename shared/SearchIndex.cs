using System.Collections;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using SQLitePCL;

namespace Windexter.Core
{
    /// TODO:
    /// - Map PropMap.db Id's to .db Metadata and .edb PropertyStore Columns
    /// - HRESULT for VT_ERROR https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/a9046ed2-bfb2-4d56-a719-2824afce59ac
    ///   and https://learn.microsoft.com/en-us/windows/win32/search/-search-prth-error-constants
    public static partial class SearchIndex
    {
        private const int ExcelMaxDataRows = 1048575;
        private const int ExcelMaxCellLength = 32767;
        private const int ExcelMaxCellsPerSheet = 10000000;
        private const string EsePropertyStoreTable = "SystemIndex_PropertyStore";
        private const string EseNameTable = "__NameTable__";
        private const uint SearchTextObfuscationMask = 0x05000113;
        private const byte SearchTextRuns = 0x00;
        private const byte SearchTextOneBytePerCharacter = 0x01;
        private const byte SearchTextCompressed = 0x02;
        private const int SearchTextCompressedHeader = 3;
        private const byte IdentifiedTextType = 0x04;
        private const int IdentifiedTextMinimumHeader = 4;
        private const byte CompressedTextHeader = 0x10;
        private const byte CompressedTextHeaderMask = 0xF8;
        private const char MapiCharacterBase = '가';
        private static readonly Regex EseLegacyPropertyTable = new(@"^SystemIndex_[0-9A-F]+A$", RegexOptions.Compiled);
        private static readonly Regex MapiIdentifier = new(@"(?<=^|[/\\]|at=)[가-곿]{2,}가(?=$|[/\\:]| \()", RegexOptions.Compiled);
        private static string dbFile = "";
        private static string dbType = "";
        private static bool eseBigEndianIntegers = false;
        private static List<List<object>> rows = [];
        private static List<List<string>> paths = [];
        private static Dictionary<int, string> resolvedPaths = [];
        private static List<List<object>> GatherResults = [];
        private static bool GatherAvailable = false;
        private static bool PropMapAvailable = false;
        private static List<List<object>> IndexResults = [];
        private static List<List<object>> URLResults = [];
        private static List<List<object>> GPSResults = [];
        private static List<List<object>> SummaryResults = [];
        private static List<List<object>> CompInfoResults = [];
        private static List<List<object>> ActivityResults = [];
        private static List<List<object>> Timeline = [];
        private static List<List<object>> PropertyMapResults = [];
        private static readonly List<string> tables = [];
        private static List<List<object>> propertyStore = [];
        private static List<List<object>> propertyMetadata = [];
        private static List<List<object>> propertyMap = [];

        public readonly record struct FilterOptions(bool Url, bool Gps, bool Summary, bool CompInfo, bool Activity, bool Timeline)
        {
            public bool Any => Url || Gps || Summary || CompInfo || Activity || Timeline;
        }

        public static string DbFile
        {
            get => dbFile;
            set => dbFile = value;
        }

        public static string OutputFileName(string timestamp, string extension = "xlsx") => $"WINDOWS-SEARCH-{dbType.ToUpper()}-{timestamp}.{extension}";

        public static Dictionary<string, List<List<object>>> GetAllResults() => new()
        {
            ["Indexed Results"] = IndexResults,
            ["Gather Data"] = GatherResults,
            ["URL Data"] = URLResults,
            ["GPS Data"] = GPSResults,
            ["Computer Info"] = CompInfoResults,
            ["Activity"] = ActivityResults,
            ["Search Summary"] = SummaryResults,
            ["Timeline"] = Timeline,
            ["Property Map"] = PropertyMapResults,
            ["Security Descriptors"] = SecurityResults,
            ["Security Entries"] = SecurityEntryResults,
        };

        public static void Reset()
        {
            rows = [];
            paths = [];
            resolvedPaths = [];
            IndexResults = [];
            GatherResults = [];
            URLResults = [];
            ActivityResults = [];
            SummaryResults = [];
            GPSResults = [];
            CompInfoResults = [];
            Timeline = [];
            propertyStore = [];
            propertyMetadata = [];
            propertyMap = [];
            PropertyMapResults = [];
            securityStore = [];
            SecurityResults = [];
            SecurityEntryResults = [];
            GatherAvailable = false;
            PropMapAvailable = false;
            SecStoreAvailable = false;
            tables.Clear();
        }

        public static string? FindGatherFile() => dbType == "index" ? FindGatherFile(dbFile) : null;

        public static string? FindGatherFile(string path)
        {
            try
            {
                string name = Path.GetFileName(path);
                if (!File.Exists(path) || name.Contains("gather", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }
                string gatherName = name.Replace("Windows.db", "Windows-gather.db", StringComparison.OrdinalIgnoreCase);
                string gatherFile = Path.Combine(Path.GetDirectoryName(path)!, gatherName);
                return gatherName != name && File.Exists(gatherFile) ? gatherFile : null;
            }
            catch (Exception ex) when (ex is ArgumentException or IOException or UnauthorizedAccessException or NotSupportedException)
            {
                return null;
            }
        }

        public static string? FindPropMapFile() => dbType == "index" ? SearchBelow(dbFile, "PropMap.db") : null;

        public static string? FindPropMapFile(string path)
        {
            string name = Path.GetFileName(path);
            return !name.EndsWith(".db", StringComparison.OrdinalIgnoreCase) || name.Contains("gather", StringComparison.OrdinalIgnoreCase) || name.Equals("PropMap.db", StringComparison.OrdinalIgnoreCase) || name.Equals(SecStoreFileName, StringComparison.OrdinalIgnoreCase) ? null : SearchBelow(path, "PropMap.db");
        }

        private static string? SearchBelow(string path, string fileName)
        {
            try
            {
                return File.Exists(path) ? Directory.EnumerateFiles(Path.GetDirectoryName(path)!, fileName, new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true }).FirstOrDefault() : null;
            }
            catch (Exception ex) when (ex is ArgumentException or IOException or UnauthorizedAccessException or NotSupportedException)
            {
                return null;
            }
        }

        public static void ExtractGather(string gatherFile)
        {
            ExtractCompanion(gatherFile);
            GatherAvailable = true;
        }

        public static void ExtractPropMap(string propMapFile)
        {
            ExtractCompanion(propMapFile);
            PropMapAvailable = true;
        }

        private static void ExtractCompanion(string companionFile)
        {
            string primaryFile = dbFile;
            string primaryType = dbType;
            dbFile = companionFile;
            try
            {
                ExtractData();
            }
            finally
            {
                dbFile = primaryFile;
                dbType = primaryType;
            }
        }

        static List<string> GetMatchingFlags(long value, Dictionary<long, string> dict)
        {
            var matches = new List<string>();

            foreach (var kvp in dict)
            {
                if ((value & kvp.Key) == kvp.Key)
                {
                    matches.Add(kvp.Value);
                }
            }

            return matches;
        }

        public static async Task ParseDataAsync()
        {
            if (dbType == "index")
            {
                if (propertyMetadata.Count == 0 || propertyStore.Count == 0)
                {
                    throw new InvalidDataException("The Windows Search property store tables contain no data.");
                }
                await Task.Run(() => GetIndexPropertyStore(propertyStore, propertyMetadata));
                if (GatherAvailable)
                {
                    ParseGatherData();
                }
                if (PropMapAvailable)
                {
                    ParsePropertyMap();
                }
                if (SecStoreAvailable)
                {
                    ParseSecurityDescriptors();
                }
            }
            else if (dbType == "esedb" || dbType == "gather")
            {
                ParseGatherData();
                if (dbType == "esedb")
                {
                    GetEsePropertyStore();
                }
            }
            else if (dbType == "propmap")
            {
                ParsePropertyMap();
            }
            else if (dbType == "secstore")
            {
                ParseSecurityDescriptors();
            }
            AddDecodedColumn(IndexResults, "System.ItemUrl", "DecodedItemUrl", IsMapiUrl);
            AddDecodedColumn(IndexResults, "System.ItemPathDisplay", "DecodedItemPathDisplay", IsMapiDisplayPath);
            AddDecodedColumn(IndexResults, "System.ItemPathDisplayNarrow", "DecodedItemPathDisplayNarrow", IsMapiDisplayPath);
            AddDecodedColumn(GatherResults, "FullPath", "DecodedFullPath", IsMapiUrl);
            AddDecodedColumn(GatherResults, "FileName", "DecodedFileName", text => MapiIdentifier.IsMatch(text));
            AddGatherCorrelation();
            AddGatherSecurity();
            AddPropertyMapNames();
            AddInvertedOnlyNames();
            rows = [];
            paths = [];
            resolvedPaths = [];
            propertyStore = [];
            propertyMetadata = [];
            propertyMap = [];
        }

        private static void ParseGatherData()
        {
            try
            {
                var resolvedPathsDict = resolvedPaths.ToDictionary(kvp => (long)kvp.Key, kvp => kvp.Value);
                List<object> header = rows[0];
                int fn = header.IndexOf("FileName");
                List<object> newHeader = ["FullPath"];
                newHeader.AddRange(header);
                GatherResults.Add(newHeader);
                for (int i = 1; i < rows.Count; i++)
                {
                    var row = rows[i];
                    if (row.Count > 0)
                    {
                        var newRow = new List<object>();
                        string fullPath = "";
                        for (int j = 0; j < row.Count; j++)
                        {
                            object item = row[j];
                            object value = item;
                            string colName = header[j].ToString()!;
                            if (colName == "ScopeID")
                            {
                                object key = dbType == "esedb" ? (uint)item : (long)item;
                                if (row[fn] is string fileName)
                                {
                                    string path = resolvedPathsDict.GetValueOrDefault((long)key, "::Unknown::/");
                                    fullPath = string.Concat(path, fileName);
                                    newRow.Insert(0, fullPath);
                                    value = key;
                                }
                            }
                            else if (colName == "UserData" && item is byte[] userDataBytes)
                            {
                                value = SPSParser.ParseToJson(userDataBytes);
                            }
                            else if (colName == "LastModified" && item is byte[] dtcellBytes)
                            {
                                if (dtcellBytes.Length == 8 && !dtcellBytes.All(b => b == 0x2A))
                                {
                                    if (dbType == "esedb")
                                    {
                                        Array.Reverse(dtcellBytes);
                                    }
                                    long filetime = BitConverter.ToInt64(dtcellBytes, 0);
                                    if (filetime != 1)
                                    {
                                        value = DateTime.FromFileTimeUtc(filetime).ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    else
                                    {
                                        value = null!;
                                    }
                                }
                                else
                                {
                                    value = null!;
                                }
                            }
                            else if (colName == "StartAddressIdentifier" && item is byte[] startAddId)
                            {
                                value = BitConverter.ToInt16(startAddId, 0);
                            }
                            else if ((colName == "Priority" || colName == "FailureUpdateAttempts") && item is byte[] singleByteArr)
                            {
                                value = singleByteArr[0];
                            }
                            else if (colName == "TransactionExtendedFlags" && value is not null && value.GetType() == typeof(uint) && (uint)value == 707406378)
                            {
                                value = null!;
                            }
                            else if (colName == "TransactionExtendedFlags" && value is not null && value.GetType() == typeof(long) && (long)value == 3038287259199220266)
                            {
                                value = null!;
                            }
                            newRow.Add(value!);
                        }
                        GatherResults.Add(newRow);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidDataException($"Unable to parse Gather Data: {ex.Message}", ex);
            }
        }

        private static void ParsePropertyMap()
        {
            List<object> header = propertyMap[0];
            PropertyMapResults.Add(header);
            try
            {
                for (int i = 1; i < propertyMap.Count; i++)
                {
                    var row = propertyMap[i];
                    if (row.Count > 0)
                    {
                        var newRow = new List<object>();
                        for (int j = 0; j < row.Count; j++)
                        {
                            object item = row[j];
                            object value = item;
                            string colName = header[j].ToString()!;
                            if (colName == "FormatId" && value is byte[] formatIdBytes && formatIdBytes.Length == 16)
                            {
                                value = new Guid(BitConverter.ToString(formatIdBytes).Replace("-", "")).ToString("B").ToUpper();
                            }
                            newRow.Add(value);
                        }
                        PropertyMapResults.Add(newRow);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidDataException($"Unable to parse Property Map: {ex.Message}", ex);
            }
        }

        private static async Task GenerateTimelineAsync()
        {
            var result = await Task.Run(() =>
            {
                var header = IndexResults[0].Cast<string>().ToList();

                var dateTimeIndexes = header
                    .Select((name, index) => (name, index))
                    .Where(x => dateTimes.Contains(x.name))
                    .ToList();

                var staticIndexes = header
                    .Select((name, index) => (name, index))
                    .Where(x => !dateTimes.Contains(x.name))
                    .ToList();
                var tempTimeline = new List<List<object>>();
                var timelineHeader = new List<object>
            {
                "Timestamp",
                "Source"
            };
                timelineHeader.AddRange(staticIndexes.Select(x => x.name));
                tempTimeline.Add(timelineHeader);
                for (int r = 1; r < IndexResults.Count; r++)
                {
                    var row = IndexResults[r];

                    foreach (var (sourceName, dtIndex) in dateTimeIndexes)
                    {
                        var value = row[dtIndex];

                        if (value is not string timestamp || timestamp.Length == 0)
                            continue;

                        var timelineRow = new List<object>
                    {
                        timestamp,
                        sourceName
                    };
                        foreach (var (_, staticIndex) in staticIndexes)
                        {
                            timelineRow.Add(row[staticIndex]);
                        }
                        tempTimeline.Add(timelineRow);
                    }
                }
                return tempTimeline;
            });
            Timeline = result;
        }

        public static async Task ParseSubData(FilterOptions filters, Action<string>? progress = null)
        {
            if (dbType == "index" || dbType == "esedb")
            {
                if (filters.Url)
                {
                    List<List<object>> temp = FilterColumns(IndexResults, UrlColumns);
                    int urlIdx = temp[0].IndexOf("System.Link.TargetUrl");
                    if (urlIdx > 0)
                    {
                        URLResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if (row[urlIdx] is string s && !string.IsNullOrEmpty(s))
                            {
                                URLResults.Add(row);
                            }
                        }
                    }
                }
                if (filters.Gps)
                {
                    List<List<object>> temp = FilterColumns(IndexResults, GpsColumns);
                    int latIdx = temp[0].IndexOf("System.GPS.LatitudeDecimal");
                    int longIdx = temp[0].IndexOf("System.GPS.LongitudeDecimal");
                    if (latIdx > 0 && longIdx > 0)
                    {
                        GPSResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if ((row[latIdx] is double lat && !double.IsNaN(lat)) || (row[longIdx] is double lon && !double.IsNaN(lon)))
                            {
                                GPSResults.Add(row);
                            }
                        }
                    }
                }
                if (filters.Summary)
                {
                    List<List<object>> temp = FilterColumns(IndexResults, SummaryColumns);
                    int summaryIdx = temp[0].IndexOf("System.Search.AutoSummary");
                    if (summaryIdx > 0)
                    {
                        SummaryResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if (row[summaryIdx] is string s && !string.IsNullOrEmpty(s))
                            {
                                SummaryResults.Add(row);
                            }
                        }
                    }
                }
                if (filters.CompInfo)
                {
                    List<List<object>> temp = FilterColumns(IndexResults, CompInfoColumns);
                    int itemTypeIdx = temp[0].IndexOf("System.ItemType");
                    if (itemTypeIdx > 0)
                    {
                        CompInfoResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if (row[itemTypeIdx] is string s && CompInfoItemTypes.Contains(s))
                            {
                                CompInfoResults.Add(row);
                            }
                        }
                    }
                }
                if (filters.Activity)
                {
                    List<List<object>> temp = FilterColumns(IndexResults, ActivityColumns);
                    int itemTypeIdx = temp[0].IndexOf("System.ItemType");
                    if (itemTypeIdx > 0)
                    {
                        ActivityResults.Add(temp[0]);
                        temp.Remove(temp[0]);
                        foreach (List<object> row in temp)
                        {
                            if (row[itemTypeIdx] is string s && s == ActivityItemType)
                            {
                                ActivityResults.Add(row);
                            }
                        }
                    }
                }
                if (filters.Timeline)
                {
                    progress?.Invoke("Generating timeline ...");
                    await GenerateTimelineAsync();
                }
            }
        }

        public static void ExtractData()
        {
            dbType = "";
            if (dbFile.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
            {
                Batteries_V2.Init();
                sqlite3 db = null!;
                sqlite3_stmt stm = null!;
                sqlite3_stmt stmt_path = null!;

                try
                {
                    int rc = raw.sqlite3_open_v2(dbFile, out db, raw.SQLITE_OPEN_READONLY, null);
                    if (rc != raw.SQLITE_OK)
                    {
                        throw new Exception($"Failed to open database: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                    }

                    rc = raw.sqlite3_create_collation(db, "UNICODE_en-US_LINGUISTIC_IGNORECASE", null,
                        (userData, a, b) => string.Compare(a, b, StringComparison.OrdinalIgnoreCase));
                    if (rc != raw.SQLITE_OK)
                    {
                        throw new Exception($"Failed to register collation: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                    }

                    string sql = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';";
                    rc = raw.sqlite3_prepare_v2(db, sql, out stm);
                    if (rc != raw.SQLITE_OK)
                    {
                        throw new Exception($"Failed to prepare SQL statement: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                    }

                    while (raw.sqlite3_step(stm) == raw.SQLITE_ROW)
                    {
                        string? tableName = raw.sqlite3_column_text(stm, 0).utf8_to_string();
                        if (tableName != null)
                        {
                            tables.Add(tableName);
                        }
                    }
                    raw.sqlite3_finalize(stm);
                    stm = null!;
                    if (tables.Contains("SystemIndex_Gthr"))
                    {
                        dbType = "gather";
                        sql = "SELECT * FROM SystemIndex_Gthr";
                        rows = ReadAndStoreTable(db, sql);

                        string sql_path = "SELECT * FROM SystemIndex_GthrPth";
                        rc = raw.sqlite3_prepare_v2(db, sql_path, out stmt_path);
                        if (rc != raw.SQLITE_OK)
                        {
                            throw new Exception($"Failed to prepare SELECT * query for SystemIndex_GthrPth: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                        }
                        while ((rc = raw.sqlite3_step(stmt_path)) == raw.SQLITE_ROW)
                        {
                            int trhColCount = raw.sqlite3_column_count(stmt_path);
                            List<string> cols = new(trhColCount);
                            for (int i = 0; i < trhColCount; i++)
                            {
                                string? text = raw.sqlite3_column_text(stmt_path, i).utf8_to_string();
                                cols.Add(text ?? string.Empty);
                            }
                            paths.Add(cols);
                        }
                        if (rc != raw.SQLITE_DONE)
                        {
                            throw new Exception($"Error iterating SystemIndex_GthrPth rows: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                        }
                        raw.sqlite3_finalize(stmt_path);
                        stmt_path = null!;
                        resolvedPaths = BuildFullPaths(paths);
                    }
                    else if (tables.Contains("CatalogStorageManager"))
                    {
                        dbType = "index";
                        string propStoreTable = tables.FirstOrDefault(t => t.EndsWith("_PropertyStore")) ?? string.Empty;
                        string metaTable = tables.FirstOrDefault(t => t.EndsWith("_Metadata")) ?? string.Empty;

                        if (string.IsNullOrEmpty(propStoreTable) || string.IsNullOrEmpty(metaTable))
                        {
                            throw new InvalidDataException("The Windows Search property store tables were not found in this database.");
                        }
                        sql = $"SELECT * FROM {propStoreTable}";
                        propertyStore = ReadAndStoreTable(db, sql);

                        sql = $"SELECT * FROM {metaTable}";
                        propertyMetadata = ReadAndStoreTable(db, sql);
                    }
                    else if (tables.Contains("PropertyMap"))
                    {
                        dbType = "propmap";
                        sql = $"SELECT * FROM PropertyMap";
                        propertyMap = ReadAndStoreTable(db, sql);
                    }
                    else if (tables.Contains(SecurityDescriptorTable))
                    {
                        dbType = "secstore";
                        sql = $"SELECT * FROM {SecurityDescriptorTable}";
                        securityStore = ReadAndStoreTable(db, sql);
                    }
                    else
                    {
                        throw new InvalidDataException("This database is not a recognised Windows Search index, gather, property map or security store database.");
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidDataException($"Unable to extract data: {ex.Message}", ex);
                }
                finally
                {
                    if (stm != null) raw.sqlite3_finalize(stm);
                    if (stmt_path != null) raw.sqlite3_finalize(stmt_path);
                    if (db != null) raw.sqlite3_close(db);
                    tables.Clear();
                }
            }
            else if (dbFile.EndsWith(".edb", StringComparison.OrdinalIgnoreCase))
            {
                dbType = "esedb";
                try
                {
                    rows.Clear();
                    using var reader = new LibEsedbReader();
                    reader.Open(dbFile);
                    var gatherPaths = ReadRequiredTable(reader, "SystemIndex_GthrPth");
                    var gatherData = ReadRequiredTable(reader, "SystemIndex_Gthr");
                    rows.Add([.. ((IEnumerable)gatherData[0]).Cast<object>()]);
                    foreach (List<object> row in (IEnumerable)gatherData.Skip(1))
                    {
                        rows.Add(row);
                    }
                    resolvedPaths = BuildEsePaths(gatherPaths);
                    var propStore = ReadEsePropertyTable(reader);
                    var propStoreHeader = (IEnumerable)propStore[0];
                    List<string> newHeader = [];
                    foreach (string title in propStoreHeader)
                    {
                        newHeader.Add((title.Contains('-') ? title.Split('-', 2)[1] : title).Replace("_", "."));
                    }
                    propertyStore.Add([.. newHeader.Cast<object>()]);
                    foreach (List<object> row in (IEnumerable)propStore.Skip(1))
                    {
                        propertyStore.Add(row);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidDataException($"Unable to extract data from ESE database: {ex.Message}", ex);
                }
            }
            else
            {
                throw new InvalidDataException("Unsupported file type. Select a Windows.db, Windows-gather.db or Windows.edb file.");
            }
        }

        private static List<object> ReadEsePropertyTable(LibEsedbReader reader)
        {
            List<object> table = reader.ReadTableData(EsePropertyStoreTable);
            eseBigEndianIntegers = false;
            if (table.Count > 0)
            {
                return table;
            }
            eseBigEndianIntegers = true;
            List<string> candidates = [.. reader.GetTableNames().Where(name => EseLegacyPropertyTable.IsMatch(name)).Order(StringComparer.Ordinal)];
            if (candidates.Count == 0)
            {
                throw new InvalidDataException($"Table {EsePropertyStoreTable} was not found in this database. Tables present: {string.Join(", ", reader.GetTableNames())}.");
            }
            List<object> nameTable = reader.ReadTableData(EseNameTable);
            string? master = null;
            if (nameTable.Count > 1 && ((IEnumerable)nameTable[0]).Cast<object>().ToList().IndexOf("CurrentMaster") is int column and >= 0 && ((IList)nameTable[1])[column] is byte[] { Length: > 0 } masterId)
            {
                master = $"SystemIndex_{masterId[0]:X}A";
            }
            return reader.ReadTableData(candidates.Contains(master!) ? master! : candidates[^1]);
        }

        private static List<object> ReadRequiredTable(LibEsedbReader reader, string tableName)
        {
            List<object> table = reader.ReadTableData(tableName);
            if (table.Count == 0)
            {
                throw new InvalidDataException($"Table {tableName} was not found in this database.");
            }
            return table;
        }

        public static Dictionary<int, string> BuildEsePaths(List<object> gatherPaths)
        {
            List<List<string>> paths = [];
            foreach (List<string> listEntry in gatherPaths.Skip(1).Cast<List<string>>())
            {
                paths.Add(listEntry);
            }
            var builtPaths = BuildFullPaths(paths);
            return builtPaths;
        }

        private static List<List<object>> ReadAndStoreTable(sqlite3 db, string sql)
        {
            var results = new List<List<object>>();
            int rc = raw.sqlite3_prepare_v2(db, sql, out sqlite3_stmt stmt);
            if (rc != raw.SQLITE_OK)
            {
                throw new Exception($"Failed to prepare SQL statement: {raw.sqlite3_errmsg(db).utf8_to_string()}");
            }

            try
            {
                int colCount = raw.sqlite3_column_count(stmt);

                var headers = new List<object>(colCount);
                for (int i = 0; i < colCount; i++)
                {
                    string? headerName = raw.sqlite3_column_name(stmt, i).utf8_to_string();
                    headers.Add(headerName ?? string.Empty);
                }
                results.Add(headers);

                while ((rc = raw.sqlite3_step(stmt)) == raw.SQLITE_ROW)
                {
                    var colValues = new List<object>(colCount);
                    for (int i = 0; i < colCount; i++)
                    {
                        int colType = raw.sqlite3_column_type(stmt, i);
                        object value;
                        switch (colType)
                        {
                            case raw.SQLITE_INTEGER:
                                value = raw.sqlite3_column_int64(stmt, i);
                                break;
                            case raw.SQLITE_FLOAT:
                                value = raw.sqlite3_column_double(stmt, i);
                                break;
                            case raw.SQLITE_TEXT:
                                value = raw.sqlite3_column_text(stmt, i).utf8_to_string();
                                break;
                            case raw.SQLITE_NULL:
                                value = DBNull.Value;
                                break;
                            case raw.SQLITE_BLOB:
                                {
                                    ReadOnlySpan<byte> blobSpan = raw.sqlite3_column_blob(stmt, i);
                                    value = blobSpan.ToArray();
                                    break;
                                }
                            default:
                                throw new NotSupportedException($"Unsupported column type {colType} at column {i}");
                        }
                        colValues.Add(value);
                    }
                    results.Add(colValues);
                }
                if (rc != raw.SQLITE_DONE)
                {
                    throw new Exception($"Error iterating rows: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidDataException($"Unable to read table data: {ex.Message}", ex);
            }
            finally
            {
                if (stmt != null)
                {
                    raw.sqlite3_finalize(stmt);
                }
            }
            return results;
        }

        private static string? DecodeConversationIndex(byte[] bytes)
        {
            if (bytes.Length < 10 || bytes.Length % 2 != 0)
            {
                return null;
            }
            uint variantType = BitConverter.ToUInt32(bytes, 0);
            uint count = BitConverter.ToUInt32(bytes, 4);
            if (variantType != 0x1F || count != (bytes.Length - 8) / 2)
            {
                return null;
            }
            if (BitConverter.ToUInt16(bytes, bytes.Length - 2) != 0)
            {
                return null;
            }
            var decoded = new byte[count - 1];
            for (int i = 0; i < decoded.Length; i++)
            {
                ushort character = BitConverter.ToUInt16(bytes, 8 + i * 2);
                if (character < MapiCharacterBase || character > MapiCharacterBase + 0xFF)
                {
                    return null;
                }
                decoded[i] = (byte)(character - MapiCharacterBase);
            }
            return Convert.ToHexString(decoded);
        }

        private static string DecodeMapiIdentifiers(string url) => MapiIdentifier.Replace(url, match => Convert.ToHexString(match.Value.Select(character => (byte)(character - MapiCharacterBase)).ToArray()));

        private static bool IsMapiUrl(string text) => text.StartsWith("mapi", StringComparison.OrdinalIgnoreCase);

        private static bool IsMapiDisplayPath(string text) => text.Contains("{S-", StringComparison.Ordinal) && MapiIdentifier.IsMatch(text);

        private static void AddDecodedColumn(List<List<object>> results, string sourceColumn, string decodedColumn, Func<string, bool> applies)
        {
            int index = results.Count > 1 ? results[0].IndexOf(sourceColumn) : -1;
            if (index < 0)
            {
                return;
            }
            var decoded = results.Skip(1).Select(row => row.Count > index && row[index] is string text && applies(text) ? DecodeMapiIdentifiers(text) : null).ToList();
            if (!results.Skip(1).Zip(decoded).Any(pair => pair.Second is not null && pair.Second != (pair.First.Count > index ? pair.First[index] as string : null)))
            {
                return;
            }
            results[0].Insert(index + 1, decodedColumn);
            for (int row = 1; row < results.Count; row++)
            {
                results[row].Insert(Math.Min(index + 1, results[row].Count), decoded[row - 1] ?? (object)DBNull.Value);
            }
        }

        private static string DecodeAutoSummary(byte[] bytes)
        {
            if (bytes.Length == 0)
            {
                return "";
            }
            uint mask = SearchTextObfuscationMask ^ (uint)bytes.Length;
            byte[] plain = [.. bytes.Select((value, index) => (byte)(value ^ (byte)(mask >> (8 * (index % 4))) ^ (byte)index))];
            byte method = plain[0];
            byte[] body = plain[1..];
            if ((method & SearchTextCompressed) != 0)
            {
                if (plain.Length < SearchTextCompressedHeader || LzxpressHuffman.Decompress(plain, SearchTextCompressedHeader, BitConverter.ToUInt16(plain, 1)) is not byte[] expanded)
                {
                    return BitConverter.ToString(bytes).Replace("-", "");
                }
                method &= unchecked((byte)~SearchTextCompressed);
                body = expanded;
            }
            StringBuilder text = new();
            if (method == SearchTextOneBytePerCharacter)
            {
                foreach (byte value in body)
                {
                    text.Append((char)value);
                }
            }
            else if (method == SearchTextRuns)
            {
                int position = 0;
                while (position < body.Length)
                {
                    if (position + 1 >= body.Length || position + 2 + body[position] > body.Length)
                    {
                        return BitConverter.ToString(bytes).Replace("-", "");
                    }
                    for (int index = 0; index < body[position]; index++)
                    {
                        text.Append((char)(body[position + 1] << 8 | body[position + 2 + index]));
                    }
                    position += 2 + body[position];
                }
            }
            else
            {
                return BitConverter.ToString(bytes).Replace("-", "");
            }
            return text.ToString().TrimEnd('\0');
        }

        private static string? DecodeIdentifiedText(byte[] bytes)
        {
            if (bytes.Length <= IdentifiedTextMinimumHeader || bytes[0] != IdentifiedTextType || bytes[1] != 0)
            {
                return null;
            }
            int header = BitConverter.ToUInt16(bytes, 2);
            if (header < IdentifiedTextMinimumHeader || header >= bytes.Length || (bytes.Length - header) % 2 != 0)
            {
                return null;
            }
            string text = Encoding.Unicode.GetString(bytes, header, bytes.Length - header);
            return text.Any(character => char.IsControl(character) || character == '�') ? null : text;
        }

        private static long ReadInt64(byte[] bytes) => eseBigEndianIntegers ? System.Buffers.Binary.BinaryPrimitives.ReadInt64BigEndian(bytes) : BitConverter.ToInt64(bytes, 0);

        private static ulong ReadUInt64(byte[] bytes) => (ulong)ReadInt64(bytes);

        private static long FutureFileTime() => (DateTime.UtcNow.AddYears(5) - new DateTime(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc)).Ticks;

        private static void GetEsePropertyStore()
        {
            try
            {
                long futureTicks = FutureFileTime();
                var header = propertyStore[0];
                IndexResults = [header];
                foreach (var entry in propertyStore.Skip(1))
                {
                    var row = new List<object> { };
                    for (int j = 0; j < entry.Count; j++)
                    {
                        object item = entry[j];
                        string colName = header[j].ToString()!;
                        object value = item;
                        if (booleanField.Contains(colName))
                        {
                            if (value is byte[] boolBytes)
                            {
                                value = BitConverter.ToBoolean(boolBytes);
                            }
                        }
                        if (value is not null && value.GetType() == typeof(long) && (long)value == 3038287259199220266)
                        {
                            value = null!;
                        }
                        else if (value is not null && value.GetType() == typeof(uint) && (uint)value == 707406378)
                        {
                            value = null!;
                        }
                        if (value is byte[] bytes)
                        {
                            if (bytes.Length == 8 && bytes.All(b => b == 0x2A))
                            {
                                value = null!;
                            }
                            else if (bytes.Length == 8 && dateTimes.Contains(colName.ToString()!))
                            {
                                try
                                {
                                    long fileTime = ReadInt64(bytes);
                                    if (fileTime <= futureTicks && fileTime > 116444736000000000)
                                    {
                                        value = DateTime.FromFileTimeUtc(fileTime).ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    else
                                    {
                                        value = "";
                                    }
                                }
                                catch
                                {
                                    value = BitConverter.ToString(bytes).Replace("-", "");
                                }
                            }
                            else if (unicodeField.Contains(colName.ToString()!))
                            {
                                value = DecodeTextValues(bytes);
                            }
                            else if (uint64List.Contains(colName.ToString()!))
                            {
                                value = ReadUInt64(bytes);
                            }
                            else if (colName.ToString()!.Equals("System.Size", StringComparison.OrdinalIgnoreCase))
                            {
                                value = ReadInt64(bytes);

                            }
                            else if (colName.ToString()!.Equals("InvertedOnlyPids", StringComparison.OrdinalIgnoreCase))
                            {
                                List<string> pids = [];
                                for (int i = 0; i < bytes.Length; i += 2)
                                {
                                    string pid = BitConverter.ToInt16(bytes, i).ToString();
                                    pids.Add(pid);

                                }
                                value = string.Join("|", pids);
                            }
                            else if (colName.ToString()!.Equals("System.Message.ConversationIndex", StringComparison.OrdinalIgnoreCase))
                            {
                                List<string> indices = [];
                                for (int i = 0; i < bytes.Length; i += 2)
                                {
                                    string index = BitConverter.ToUInt16(bytes, i).ToString();
                                    indices.Add(index);
                                }
                                value = DecodeConversationIndex(bytes) ?? string.Join("|", indices);
                            }
                            else if (durations.Contains(colName.ToString()!))
                            {
                                ulong time = ReadUInt64(bytes);
                                TimeSpan duration = TimeSpan.FromTicks((long)time);
                                value = string.Format("{0:D2}:{1:D2}:{2:D2}:{3:D2}", (int)duration.Days, (int)duration.Hours, duration.Minutes, duration.Seconds);
                                var description = GetMeasurementName(time, MediaDurations);
                                value = $"{value} - {description}";
                            }
                            else if (colName.ToString()!.Equals("System.Search.AutoSummary", StringComparison.OrdinalIgnoreCase))
                            {
                                value = DecodeAutoSummary(bytes);
                            }
                            else if (floatField.Contains(colName.ToString()!))
                            {
                                value = BitConverter.ToDouble(bytes);
                            }
                            else if (colName.ToString()!.Equals("System.GPS.Longitude") | colName.ToString()!.Equals("System.GPS.Latitude"))
                            {
                                value = FormatGpsCoordinate(bytes);
                            }
                            else if (guids.Contains(colName.ToString()!))
                            {
                                value = BitConverter.ToString(bytes).Replace("-", "");
                                Guid guid = new((string)value);
                                value = guid.ToString("B").ToUpper();
                            }
                            else if (lookups.Contains(colName.ToString()!) && value is not null)
                            {
                                if (LookupValues.TryGetValue(colName.ToString()!, out var lookup))
                                {
                                    value = BitConverter.ToInt16(bytes, 0);
                                    if (lookup.TryGetValue((short)value, out string? strValue))
                                    {
                                        value = strValue;
                                    }
                                }
                            }
                            else
                            {
                                value = DecodeIdentifiedText(bytes) ?? BitConverter.ToString(bytes).Replace("-", "");
                            }
                        }
                        if (sfgaoField.Contains(colName.ToString()!) && value is not null)
                        {
                            List<string> flags = GetMatchingFlags((uint)value, SFGAO);
                            value = string.Join("|", flags);
                        }
                        if (colName.ToString()!.Equals("System.Activity.BackgroundColor", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            uint colorVal = (uint)value;
                            Color c = Color.FromArgb((int)colorVal);
                            value = $"{value} - ARGB({c.A},{c.R},{c.G},{c.B})";
                        }
                        if (colName.ToString()!.Equals("System.FilePlaceholderStatus", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            List<string> matches = GetMatchingFlags((uint)value, FilePlaceholderStates);
                            value = string.Join("|", matches);
                        }
                        if (colName.ToString()!.Equals("System.FileAttributes", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            List<string> flags = GetMatchingFlags((uint)value, FileAttributes);
                            value = string.Join("|", flags);
                        }
                        if (colName.ToString()!.Equals("System.Capacity", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            value = FormatCapacity(value);
                        }
                        if (colName.ToString()!.Equals("System.Message.Flags", StringComparison.OrdinalIgnoreCase) && value is not null)
                        {
                            List<string> matches = GetMatchingFlags((uint)value, MessageFlags);
                            value = string.Join("|", matches);
                        }
                        row.Add(value!);
                    }
                    IndexResults.Add(row);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidDataException($"Unable to read and correlate data: {ex.Message}", ex);
            }
        }

        private static string? TryDecodeUtf16(byte[] bytes)
        {
            if (bytes.Length % 2 != 0)
            {
                return null;
            }
            try
            {
                string text = new UnicodeEncoding(false, false, true).GetString(bytes);
                return text.Any(character => character < ' ' && character is not ('\0' or '\t' or '\n' or '\r')) ? null : text;
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        private static string? TryDecode7Bit(byte[] bytes)
        {
            if (bytes.Length < 2 || (bytes[0] & CompressedTextHeaderMask) != CompressedTextHeader)
            {
                return null;
            }
            StringBuilder text = new();
            int accumulator = 0;
            int bits = 0;
            foreach (byte value in bytes.Skip(1))
            {
                accumulator |= value << bits;
                bits += 8;
                for (; bits >= 7; bits -= 7, accumulator >>= 7)
                {
                    char character = (char)(accumulator & 0x7F);
                    if (character != '\0' && (character < ' ' || character > '~'))
                    {
                        return null;
                    }
                    text.Append(character);
                }
            }
            return text.ToString();
        }

        private static string DecodeTextValues(byte[] bytes)
        {
            string text = TryDecode7Bit(bytes) ?? TryDecodeUtf16(bytes) ?? (dbType == "esedb" && bytes.Length > 1 && bytes[1] != 0 ? Decompress7Bit(bytes) : Encoding.Unicode.GetString(bytes));
            return string.Join("; ", text.Split('\0', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
        }

        public static string Decompress7Bit(byte[] data)
        /// https://www.3gpp.org/ftp/Specs/archive/23_series/23.038 Para 6, SMS Packing
        /// https://doubleblak.com/blogPost.php?k=7bitpdu
        {
            if (data == null || data.Length <= 1)
                return string.Empty;
            StringBuilder sb = new();
            int shift = 0;
            int prevByte = 0;
            for (int i = 1; i < data.Length; i++)
            {
                byte currentByte = data[i];
                int charCode = ((currentByte << shift) & 0x7F) | (prevByte >> (8 - shift));
                if (dbType == "esedb")
                {
                    AppendIfPrintable(sb, charCode);
                }
                else
                {
                    sb.Append((char)charCode);
                }

                shift++;
                if (shift == 7)
                {
                    // After every 7 bytes processed, there are 7 "leftover" bits in the current byte.
                    // This forms an 8th character.
                    int eighthChar = (currentByte >> 1) & 0x7F;
                    if (dbType == "esedb")
                    {
                        AppendIfPrintable(sb, eighthChar);
                    }
                    else
                    {
                        sb.Append((char)eighthChar);
                    }
                    shift = 0;
                    prevByte = 0;
                }
                else
                {
                    prevByte = currentByte;
                }
            }
            return sb.ToString().Replace('\0', ' ').Trim();
        }

        static void AppendIfPrintable(StringBuilder sb, int code)
        {
            if (code == 0)
            {
                sb.Append(' ');
            }
            else if (code >= 0x20 && code <= 0x7E)
            {
                sb.Append((char)code);
            }
        }

        private static void GetIndexPropertyStore(List<List<object>> propertyStore, List<List<object>> propertyMetadata)
        {
            try
            {
                long futureTicks = FutureFileTime();
                var idToName = propertyMetadata.Skip(1).ToDictionary(row => (long)row[0], row => (string)row[2]);
                var nameToIds = idToName.GroupBy(kvp => kvp.Value).ToDictionary(group => group.Key, group => group.Select(kvp => kvp.Key).ToArray());
                var pivotedData = propertyStore.Skip(1).GroupBy(row => (long)row[0]).ToDictionary(group => group.Key, group => group.ToDictionary(item => (long)item[1], item => item[2]));
                var workIds = pivotedData.Keys.ToList();
                var uniquePropIds = propertyStore.Skip(1).Select(row => (long)row[1]).Distinct().ToList();
                var header = new List<object> { "WorkId" };
                var seenPropNames = new HashSet<string>();
                foreach (var propId in uniquePropIds)
                {
                    if (idToName.TryGetValue(propId, out var propName) && !seenPropNames.Contains(propName))
                    {
                        header.Add(propName);
                        seenPropNames.Add(propName);
                    }
                }
                IndexResults = [header];
                foreach (var workId in workIds)
                {
                    var row = new List<object> { workId };
                    var workIdProperties = pivotedData[workId];
                    foreach (var colName in header.Skip(1))
                    {
                        var colId = nameToIds[colName.ToString()!].FirstOrDefault(id => workIdProperties.ContainsKey(id));
                        if (colId != 0 && workIdProperties.TryGetValue(colId, out object? value))
                        {
                            if (booleanField.Contains(colName.ToString()!))
                            {
                                value = Convert.ToBoolean(value);
                            }
                            if (value is not null && value.GetType() == typeof(long) && (long)value == 3038287259199220266)
                            {
                                value = null!;
                            }
                            else if (value is not null && value.GetType() == typeof(uint) && (uint)value == 707406378)
                            {
                                value = null!;
                            }
                            if (value is byte[] bytes)
                            {
                                if (bytes.Length == 8 && bytes.All(b => b == 0x2A))
                                {
                                    value = null!;
                                }
                                else if (bytes.Length == 8 && dateTimes.Contains(colName.ToString()!))
                                {
                                    try
                                    {
                                        long fileTime = BitConverter.ToInt64(bytes, 0);
                                        if (fileTime <= futureTicks && fileTime > 116444736000000000)
                                        {
                                            value = DateTime.FromFileTimeUtc(fileTime).ToString("yyyy-MM-dd HH:mm:ss");
                                        }
                                        else
                                        {
                                            value = "";
                                        }
                                    }
                                    catch
                                    {
                                        value = BitConverter.ToString(bytes).Replace("-", "");
                                    }
                                }
                                else if (unicodeField.Contains(colName.ToString()!))
                                {
                                    value = DecodeTextValues(bytes);
                                }
                                else if (uint64List.Contains(colName.ToString()!))
                                {
                                    value = BitConverter.ToUInt64(bytes);
                                }
                                else if (colName.ToString()!.Equals("System.Size", StringComparison.OrdinalIgnoreCase))
                                {
                                    long fileSize = BitConverter.ToInt64(bytes, 0);
                                    value = fileSize;
                                }
                                else if (colName.ToString()!.Equals("InvertedOnlyPids", StringComparison.OrdinalIgnoreCase))
                                {
                                    List<string> pids = [];
                                    for (int i = 0; i < bytes.Length; i += 2)
                                    {
                                        string pid = BitConverter.ToInt16(bytes, i).ToString();
                                        pids.Add(pid);
                                    }
                                    value = string.Join("|", pids);
                                }
                                else if (colName.ToString()!.Equals("System.Message.ConversationIndex", StringComparison.OrdinalIgnoreCase))
                                {
                                    List<string> indices = [];
                                    for (int i = 0; i < bytes.Length; i += 2)
                                    {
                                        string index = BitConverter.ToUInt16(bytes, i).ToString();
                                        indices.Add(index);
                                    }
                                    value = DecodeConversationIndex(bytes) ?? string.Join("|", indices);
                                }
                                else if (durations.Contains(colName.ToString()!))
                                {
                                    ulong time = BitConverter.ToUInt64((byte[])value, 0);
                                    TimeSpan duration = TimeSpan.FromTicks((long)time);
                                    value = string.Format("{0:D2}:{1:D2}:{2:D2}:{3:D2}", (int)duration.Days, (int)duration.Hours, duration.Minutes, duration.Seconds);
                                    var description = GetMeasurementName(time, MediaDurations);
                                    value = $"{value} - {description}";
                                }
                                else if (colName.ToString()!.Equals("System.GPS.Longitude") | colName.ToString()!.Equals("System.GPS.Latitude"))
                                {
                                    value = FormatGpsCoordinate(bytes);
                                }
                                else if (colName.ToString()!.Equals("System.Capacity", StringComparison.OrdinalIgnoreCase) && bytes.Length == 8)
                                {
                                    value = BitConverter.ToUInt64(bytes);
                                }
                                else
                                {
                                    value = BitConverter.ToString(bytes).Replace("-", "");
                                }
                            }
                            if (sfgaoField.Contains(colName.ToString()!) && value is not null)
                            {
                                List<string> flags = GetMatchingFlags((long)value, SFGAO);
                                value = string.Join("|", flags);
                            }
                            if (colName.ToString()!.Equals("System.Activity.BackgroundColor", StringComparison.OrdinalIgnoreCase) && value is not null)
                            {
                                uint colorVal = (uint)(long)value;
                                Color c = Color.FromArgb((int)colorVal);
                                value = $"{value} - ARGB({c.A},{c.R},{c.G},{c.B})";
                            }
                            if (colName.ToString()!.Equals("System.FilePlaceholderStatus", StringComparison.OrdinalIgnoreCase) && value is not null)
                            {
                                List<string> matches = GetMatchingFlags((long)value, FilePlaceholderStates);
                                value = string.Join("|", matches);
                            }
                            if (colName.ToString()!.Equals("System.FileAttributes", StringComparison.OrdinalIgnoreCase) && value is not null)
                            {
                                List<string> flags = GetMatchingFlags((long)value, FileAttributes);
                                value = string.Join("|", flags);
                            }
                            if (guids.Contains(colName.ToString()!) && value is not null)
                            {
                                Guid guid = new((string)value);
                                value = guid.ToString("B").ToUpper();
                            }
                            if (lookups.Contains(colName.ToString()!) && value is not null)
                            {
                                if (LookupValues.TryGetValue(colName.ToString()!, out var lookup))
                                {
                                    if (lookup.TryGetValue((int)(long)value, out string? strValue))
                                    {
                                        value = strValue;
                                    }
                                }
                            }
                            if (colName.ToString()!.Equals("System.Capacity", StringComparison.OrdinalIgnoreCase) && value is not null)
                            {
                                value = FormatCapacity(value);
                            }
                            if (colName.ToString()!.Equals("System.Message.Flags", StringComparison.OrdinalIgnoreCase) && value is not null)
                            {
                                List<string> matches = GetMatchingFlags((long)value, MessageFlags);
                                value = string.Join("|", matches);
                            }
                            row.Add(value!);
                        }
                        else
                        {
                            row.Add(DBNull.Value);
                        }
                    }
                    IndexResults.Add(row);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidDataException($"Unable to read and correlate data: {ex.Message}", ex);
            }
        }

        private static Dictionary<int, string> BuildFullPaths(List<List<string>> paths)
        {
            var lookup = new Dictionary<int, (int Parent, string Path)>();
            foreach (var row in paths)
            {
                if (row.Count >= 3 && int.TryParse(row[0], out int scopeId) && int.TryParse(row[1], out int parentId))
                {
                    lookup[scopeId] = (parentId, row[2]);
                }
            }
            var result = new Dictionary<int, string>();
            string ResolvePathIterative(int scopeId)
            {
                if (result.TryGetValue(scopeId, out string? fullPath))
                {
                    return fullPath;
                }
                if (!lookup.ContainsKey(scopeId))
                {
                    return "";
                }
                var pathParts = new Stack<string>();
                var currentId = scopeId;
                var visited = new HashSet<int>();
                while (visited.Add(currentId) && lookup.TryGetValue(currentId, out var entry))
                {
                    pathParts.Push(entry.Path);
                    if (entry.Parent == 1)
                    {
                        break;
                    }
                    currentId = entry.Parent;
                }
                var sb = new StringBuilder();
                while (pathParts.Count > 0)
                {
                    sb.Append(pathParts.Pop());
                }
                fullPath = sb.ToString();
                result[scopeId] = fullPath;
                return fullPath;
            }
            foreach (var scopeId in lookup.Keys)
            {
                ResolvePathIterative(scopeId);
            }
            return result;
        }

        public static async Task ExportToExcel(string filePath)
        {
            try
            {
                await Task.Run(() => {
                    using var workbook = new XLWorkbook();
                    foreach (var (title, results) in GetAllResults())
                    {
                        if (results is not [])
                        {
                            int start = 1, part = 1;
                            do
                            {
                                int end = NextSheetEnd(results, start);
                                WriteWorksheet(workbook, part == 1 ? title : $"{title} ({part})", results, start, end);
                                start = end;
                                part++;
                            } while (start < results.Count);
                        }
                    }
                    workbook.SaveAs(filePath);
                });
            }
            catch (Exception ex)
            {
                throw new IOException($"An error occurred while exporting: {ex.Message}", ex);
            }
        }

        private static int NextSheetEnd(List<List<object>> results, int start)
        {
            long cells = 0;
            int end = start;
            while (end < results.Count && end - start < ExcelMaxDataRows)
            {
                long rowCells = results[end].Count(value => value is not (null or DBNull or ""));
                if (end > start && cells + rowCells > ExcelMaxCellsPerSheet)
                {
                    break;
                }
                cells += rowCells;
                end++;
            }
            return end;
        }

        private static void WriteWorksheet(XLWorkbook workbook, string sheetName, List<List<object>> results, int start, int end)
        {
            var worksheet = workbook.Worksheets.Add(sheetName);
            int lastCol = results[0].Count;
            HashSet<int> dateTimeCols = [];
            for (int col = 0; col < lastCol; col++)
            {
                string header = results[0][col]?.ToString() ?? string.Empty;
                int excelCol = col + 1;

                worksheet.Cell(1, excelCol).Value = header;

                if (dateTimes.Contains(header) || header == "Timestamp")
                {
                    dateTimeCols.Add(excelCol);
                    worksheet.Column(excelCol).Style.NumberFormat.Format = "yyyy-MM-dd HH:mm:ss";
                }
                else
                {
                    worksheet.Column(excelCol).Style.NumberFormat.Format = "@";
                }
            }
            for (int row = start; row < end; row++)
            {
                var dataRow = results[row];

                for (int col = 0; col < dataRow.Count; col++)
                {
                    int excelCol = col + 1;

                    if (dateTimeCols.Contains(excelCol))
                    {
                        if (dataRow[col] is DateTime dt)
                        {
                            worksheet.Cell(row - start + 2, excelCol).Value = dt;
                        }
                        else if (DateTime.TryParse(dataRow[col]?.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
                        {
                            worksheet.Cell(row - start + 2, excelCol).Value = parsed;
                        }
                    }
                    else if (FormatValueForExcel(dataRow[col]) is { Length: > 0 } text)
                    {
                        worksheet.Cell(row - start + 2, excelCol).Value = text;
                    }
                }
            }
            worksheet.Columns().AdjustToContents(1, 1000);
            worksheet.SheetView.Freeze(1, 0);
            worksheet.Range(1, 1, end - start + 1, lastCol).SetAutoFilter();
        }

        private static string FormatValueForExcel(object value)
        {
            if (value == null || value is DBNull)
            {
                return string.Empty;
            }

            string text = value is byte[] bytes ? BitConverter.ToString(bytes).Replace("-", "") : value.ToString() ?? string.Empty;
            return text.Length > ExcelMaxCellLength ? text[..ExcelMaxCellLength] : text;
        }

        private static List<List<object>> FilterColumns(List<List<object>> source, List<string> desiredColumns)
        {
            if (source == null || source.Count == 0)
                return [];

            var headerRow = source[0];
            var headers = headerRow.Select(h => h?.ToString() ?? string.Empty).ToList();

            var selectedIndexes = desiredColumns
                .Select(col => headers.IndexOf(col))
                .Where(index => index >= 0)
                .ToList();

            return [.. source
                .Select(row => selectedIndexes
                    .Where(i => i < row.Count)
                    .Select(i => row[i])
                    .ToList())];
        }
    }
}
