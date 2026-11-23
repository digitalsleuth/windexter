using System.Globalization;
using System.IO;
using SQLitePCL;

namespace Windexter.Core
{
    public static partial class SearchIndex
    {
        private const string SqliteIndexTable = "IndexedResults";
        private const string SqliteGatherTable = "GatherData";
        private const string SqlitePropertyMapTable = "PropertyMap";
        private const string SqliteSecurityDescriptorTable = "SecurityDescriptors";
        private const string SqliteSecurityEntryTable = "SecurityEntries";
        private const string SqliteInfoTable = "ExportInfo";
        private const string SqliteTimestampFormat = "yyyy-MM-dd HH:mm:ss";

        [Flags]
        private enum SqliteColumnKind
        {
            None = 0,
            Integer = 1,
            Real = 2,
            Text = 4
        }

        public static async Task ExportToSqlite(string filePath, string toolVersion)
        {
            try
            {
                await Task.Run(() => WriteSqlite(filePath, toolVersion));
            }
            catch (Exception ex)
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (IOException)
                {
                }
                throw new IOException($"An error occurred while exporting to SQLite: {ex.Message}", ex);
            }
        }

        private static void WriteSqlite(string filePath, string toolVersion)
        {
            Batteries_V2.Init();
            File.Delete(filePath);
            sqlite3 db = null!;
            try
            {
                if (raw.sqlite3_open_v2(filePath, out db, raw.SQLITE_OPEN_READWRITE | raw.SQLITE_OPEN_CREATE, null) != raw.SQLITE_OK)
                {
                    throw new IOException($"Failed to create database: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                }
                SqliteExecute(db, "PRAGMA journal_mode=OFF; PRAGMA synchronous=OFF; PRAGMA temp_store=MEMORY; BEGIN");
                string[] indexColumns = IndexResults.Count > 0 ? WriteSqliteTable(db, SqliteIndexTable, IndexResults) : [];
                if (GatherResults.Count > 0)
                {
                    WriteSqliteTable(db, SqliteGatherTable, GatherResults);
                }
                if (PropertyMapResults.Count > 0)
                {
                    WriteSqliteTable(db, SqlitePropertyMapTable, PropertyMapResults);
                }
                if (SecurityResults.Count > 0)
                {
                    WriteSqliteTable(db, SqliteSecurityDescriptorTable, SecurityResults);
                    WriteSqliteTable(db, SqliteSecurityEntryTable, SecurityEntryResults);
                }
                if (indexColumns.Length > 0)
                {
                    CreateSqliteViews(db, IndexResults[0], indexColumns);
                }
                WriteSqliteInfo(db, toolVersion);
                SqliteExecute(db, "COMMIT");
            }
            finally
            {
                if (db != null)
                {
                    raw.sqlite3_close(db);
                }
            }
        }

        private static string[] WriteSqliteTable(sqlite3 db, string table, List<List<object>> results)
        {
            string[] names = UniqueColumnNames(results[0]);
            var kinds = new SqliteColumnKind[names.Length];
            for (int row = 1; row < results.Count; row++)
            {
                List<object> values = results[row];
                for (int col = 0; col < Math.Min(names.Length, values.Count); col++)
                {
                    kinds[col] |= SqliteKindOf(values[col]);
                }
            }
            SqliteExecute(db, $"CREATE TABLE {Quote(table)} ({string.Join(", ", names.Select((name, col) => $"{Quote(name)} {SqliteDeclaredType(kinds[col])}"))})");
            if (names.Contains("WorkId"))
            {
                SqliteExecute(db, $"CREATE INDEX {Quote($"idx_{table}_WorkId")} ON {Quote(table)} ({Quote("WorkId")})");
            }
            if (raw.sqlite3_prepare_v2(db, $"INSERT INTO {Quote(table)} VALUES ({string.Join(", ", names.Select((_, col) => $"?{col + 1}"))})", out sqlite3_stmt insert) != raw.SQLITE_OK)
            {
                throw new IOException($"Failed to prepare insert for {table}: {raw.sqlite3_errmsg(db).utf8_to_string()}");
            }
            try
            {
                for (int row = 1; row < results.Count; row++)
                {
                    List<object> values = results[row];
                    for (int col = 0; col < names.Length; col++)
                    {
                        SqliteBind(insert, col + 1, col < values.Count ? values[col] : null);
                    }
                    if (raw.sqlite3_step(insert) != raw.SQLITE_DONE)
                    {
                        throw new IOException($"Failed to insert into {table}: {raw.sqlite3_errmsg(db).utf8_to_string()}");
                    }
                    raw.sqlite3_reset(insert);
                }
            }
            finally
            {
                raw.sqlite3_finalize(insert);
            }
            return names;
        }

        private static void CreateSqliteViews(sqlite3 db, List<object> header, string[] names)
        {
            var columnNames = new Dictionary<string, string>();
            for (int col = 0; col < names.Length; col++)
            {
                columnNames.TryAdd(header[col]?.ToString() ?? string.Empty, names[col]);
            }
            void CreateFilterView(string view, List<string> columns, string[] keys, Func<string[], string> condition)
            {
                if (keys.All(columnNames.ContainsKey))
                {
                    SqliteExecute(db, $"CREATE VIEW {Quote(view)} AS SELECT {string.Join(", ", columns.Where(columnNames.ContainsKey).Select(column => Quote(columnNames[column])))} FROM {Quote(SqliteIndexTable)} WHERE {condition([.. keys.Select(key => Quote(columnNames[key]))])}");
                }
            }
            CreateFilterView("URLData", UrlColumns, ["System.Link.TargetUrl"], keys => $"{keys[0]} IS NOT NULL AND {keys[0]} <> ''");
            CreateFilterView("GPSData", GpsColumns, ["System.GPS.LatitudeDecimal", "System.GPS.LongitudeDecimal"], keys => $"{keys[0]} IS NOT NULL OR {keys[1]} IS NOT NULL");
            CreateFilterView("SearchSummary", SummaryColumns, ["System.Search.AutoSummary"], keys => $"{keys[0]} IS NOT NULL AND {keys[0]} <> ''");
            CreateFilterView("ComputerInfo", CompInfoColumns, ["System.ItemType"], keys => $"{keys[0]} IN ({string.Join(", ", CompInfoItemTypes.Select(type => $"'{type}'"))})");
            CreateFilterView("Activity", ActivityColumns, ["System.ItemType"], keys => $"{keys[0]} = '{ActivityItemType}'");
            List<string> dateColumns = [.. header.Select(column => column?.ToString() ?? string.Empty).Where(column => dateTimes.Contains(column))];
            if (dateColumns.Count > 0)
            {
                string otherColumns = string.Join(", ", header.Select(column => column?.ToString() ?? string.Empty).Where(column => !dateTimes.Contains(column)).Select(column => Quote(columnNames[column])));
                SqliteExecute(db, $"CREATE VIEW {Quote("Timeline")} AS " + string.Join(" UNION ALL ", dateColumns.Select(column => $"SELECT {Quote(columnNames[column])} AS {Quote("Timestamp")}, '{column.Replace("'", "''")}' AS {Quote("Source")}, {otherColumns} FROM {Quote(SqliteIndexTable)} WHERE {Quote(columnNames[column])} IS NOT NULL AND {Quote(columnNames[column])} <> ''")));
            }
        }

        private static void WriteSqliteInfo(sqlite3 db, string toolVersion)
        {
            SqliteExecute(db, $"CREATE TABLE {Quote(SqliteInfoTable)} ({Quote("Key")} TEXT, {Quote("Value")} TEXT)");
            if (raw.sqlite3_prepare_v2(db, $"INSERT INTO {Quote(SqliteInfoTable)} VALUES (?1, ?2)", out sqlite3_stmt insert) != raw.SQLITE_OK)
            {
                throw new IOException($"Failed to prepare insert for {SqliteInfoTable}: {raw.sqlite3_errmsg(db).utf8_to_string()}");
            }
            try
            {
                foreach (var (key, value) in new (string, string)[]
                {
                    ("Tool", toolVersion),
                    ("ExportedUtc", DateTime.UtcNow.ToString(SqliteTimestampFormat, CultureInfo.InvariantCulture)),
                    ("SourceFile", dbFile),
                    ("SourceType", dbType),
                    ("GatherIncluded", (GatherResults.Count > 0).ToString()),
                    ("PropertyMapIncluded", (PropertyMapResults.Count > 0).ToString()),
                    ("SecurityStoreIncluded", (SecurityResults.Count > 0).ToString())
                })
                {
                    raw.sqlite3_bind_text(insert, 1, key);
                    raw.sqlite3_bind_text(insert, 2, value);
                    raw.sqlite3_step(insert);
                    raw.sqlite3_reset(insert);
                }
            }
            finally
            {
                raw.sqlite3_finalize(insert);
            }
        }

        private static string[] UniqueColumnNames(List<object> header)
        {
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return [.. header.Select((column, index) =>
            {
                string baseName = string.IsNullOrWhiteSpace(column?.ToString()) ? $"Column{index + 1}" : column!.ToString()!;
                string name = baseName;
                for (int suffix = 2; !used.Add(name); suffix++)
                {
                    name = $"{baseName}_{suffix}";
                }
                return name;
            })];
        }

        private static string Quote(string identifier) => $"\"{identifier.Replace("\"", "\"\"")}\"";

        private static void SqliteExecute(sqlite3 db, string sql)
        {
            if (raw.sqlite3_exec(db, sql) != raw.SQLITE_OK)
            {
                throw new IOException($"{raw.sqlite3_errmsg(db).utf8_to_string()} ({sql[..Math.Min(sql.Length, 120)]})");
            }
        }

        private static SqliteColumnKind SqliteKindOf(object? value) => value switch
        {
            null or DBNull => SqliteColumnKind.None,
            bool or sbyte or byte or short or ushort or int or uint or long => SqliteColumnKind.Integer,
            ulong unsigned => unsigned <= long.MaxValue ? SqliteColumnKind.Integer : SqliteColumnKind.Text,
            double real => double.IsNaN(real) ? SqliteColumnKind.None : SqliteColumnKind.Real,
            float real => float.IsNaN(real) ? SqliteColumnKind.None : SqliteColumnKind.Real,
            _ => SqliteColumnKind.Text
        };

        private static string SqliteDeclaredType(SqliteColumnKind kind) => kind.HasFlag(SqliteColumnKind.Text) ? "TEXT" : kind.HasFlag(SqliteColumnKind.Real) ? "REAL" : kind.HasFlag(SqliteColumnKind.Integer) ? "INTEGER" : "TEXT";

        private static void SqliteBind(sqlite3_stmt statement, int index, object? value)
        {
            switch (value)
            {
                case null or DBNull:
                    raw.sqlite3_bind_null(statement, index);
                    break;
                case bool flag:
                    raw.sqlite3_bind_int64(statement, index, flag ? 1 : 0);
                    break;
                case sbyte or byte or short or ushort or int or uint or long:
                    raw.sqlite3_bind_int64(statement, index, Convert.ToInt64(value, CultureInfo.InvariantCulture));
                    break;
                case ulong unsigned when unsigned <= long.MaxValue:
                    raw.sqlite3_bind_int64(statement, index, (long)unsigned);
                    break;
                case double real:
                    if (double.IsNaN(real)) raw.sqlite3_bind_null(statement, index); else raw.sqlite3_bind_double(statement, index, real);
                    break;
                case float real:
                    if (float.IsNaN(real)) raw.sqlite3_bind_null(statement, index); else raw.sqlite3_bind_double(statement, index, real);
                    break;
                case byte[] bytes:
                    raw.sqlite3_bind_text(statement, index, Convert.ToHexString(bytes));
                    break;
                case DateTime timestamp:
                    raw.sqlite3_bind_text(statement, index, timestamp.ToString(SqliteTimestampFormat, CultureInfo.InvariantCulture));
                    break;
                default:
                    raw.sqlite3_bind_text(statement, index, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                    break;
            }
        }
    }
}
