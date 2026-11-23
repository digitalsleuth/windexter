using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Windexter.Core
{
    public class LibEsedbReader : IDisposable
    {
        private IntPtr _fileHandle = IntPtr.Zero;
        private bool _isOpen = false;

        #region P/Invoke Declarations

        // File operations
        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int libesedb_file_initialize(out IntPtr file, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int libesedb_file_free(ref IntPtr file, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int libesedb_file_open(IntPtr file, byte[] filename, int access_flags, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
        private static extern int libesedb_file_open_wide(IntPtr file, string filename, int access_flags, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_file_close(IntPtr file, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_file_get_number_of_tables(IntPtr file, out int number_of_tables, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_file_get_table(IntPtr file, int table_entry, out IntPtr table, out IntPtr error);

        // Table operations
        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_free(ref IntPtr table, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_utf8_name_size(IntPtr table, out UIntPtr utf8_name_size, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_utf8_name(IntPtr table, byte[] utf8_name, UIntPtr utf8_name_size, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_number_of_columns(IntPtr table, out int number_of_columns, int flags, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_column(IntPtr table, int column_entry, out IntPtr column, int flags, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_number_of_records(IntPtr table, out int number_of_records, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_table_get_record(IntPtr table, int record_entry, out IntPtr record, out IntPtr error);

        // Column operations
        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_column_free(ref IntPtr column, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_column_get_utf8_name_size(IntPtr column, out UIntPtr utf8_name_size, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_column_get_utf8_name(IntPtr column, byte[] utf8_name, UIntPtr utf8_name_size, out IntPtr error);

        // Record operations
        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_free(ref IntPtr record, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_number_of_values(IntPtr record, out int number_of_values, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_column_type(IntPtr record, int value_entry, out uint column_type, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_utf8_string_size(IntPtr record, int value_entry, out UIntPtr utf8_string_size, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_utf8_string(IntPtr record, int value_entry, byte[] utf8_value, UIntPtr utf8_value_size, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_data_size(IntPtr record, int value_entry, out UIntPtr value_data_size, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_data(IntPtr record, int value_entry, byte[] value_data, UIntPtr value_data_size, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_32bit(IntPtr record, int value_entry, out uint value_32bit, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_value_64bit(IntPtr record, int value_entry, out ulong value_64bit, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_record_get_multi_value(IntPtr record, int value_entry, out IntPtr multi_value, out IntPtr error);

        // Multi-value operations
        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_free(ref IntPtr multi_value, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_number_of_values(IntPtr multi_value, out int number_of_values, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_value_utf8_string_size(IntPtr multi_value, int multi_value_index, out UIntPtr utf8_string_size, out IntPtr error);

        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_multi_value_get_value_utf8_string(IntPtr multi_value, int multi_value_index, byte[] utf8_string, UIntPtr utf8_string_size, out IntPtr error);

        // Error handling
        [DllImport("libesedb", CallingConvention = CallingConvention.Cdecl)]
        private static extern int libesedb_error_free(ref IntPtr error);

        #endregion

        public void Open(string databasePath)
        {
            if (_isOpen) return;

            databasePath = Path.GetFullPath(databasePath);
            int result = Check(libesedb_file_initialize(out _fileHandle, out nint error), ref error);
            if (result != 1)
            {
                throw new Exception("Failed to initialize libesedb file");
            }
            try
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    result = Check(libesedb_file_open_wide(_fileHandle, databasePath, 0x01, out error), ref error);
                }
                else
                {
                    byte[] utf8Path = Encoding.UTF8.GetBytes(databasePath + '\0');
                    result = Check(libesedb_file_open(_fileHandle, utf8Path, 0x01, out error), ref error);
                }
                if (result != 1)
                {
                    throw new Exception($"Failed to open database: {databasePath}");
                }
                _isOpen = true;
            }
            catch
            {
                if (_fileHandle != nint.Zero)
                {
                    _ = Check(libesedb_file_free(ref _fileHandle, out error), ref error);
                }
                throw;
            }
        }

        public List<string> GetTableNames()
        {
            EnsureOpen();
            List<string> names = [];
            if (Check(libesedb_file_get_number_of_tables(_fileHandle, out int numberOfTables, out nint error), ref error) != 1)
            {
                return names;
            }
            for (int i = 0; i < numberOfTables; i++)
            {
                if (Check(libesedb_file_get_table(_fileHandle, i, out nint tableHandle, out error), ref error) == 1 && tableHandle != IntPtr.Zero)
                {
                    names.Add(GetTableName(tableHandle) ?? $"Table{i}");
                    _ = Check(libesedb_table_free(ref tableHandle, out error), ref error);
                }
            }
            return names;
        }

        public List<object> ReadTableData(string tableName)
        {
            EnsureOpen();
            var rows = new List<object>();
            IntPtr tableHandle = GetTableHandle(tableName);

            if (tableHandle == IntPtr.Zero)
            {
                return rows;
            }
            try
            {
                var columnInfo = new List<string>();
                int result = Check(libesedb_table_get_number_of_columns(tableHandle, out int numberOfColumns, 0, out nint error), ref error);

                if (result == 1)
                {
                    for (int i = 0; i < numberOfColumns; i++)
                    {
                        result = Check(libesedb_table_get_column(tableHandle, i, out nint columnHandle, 0, out error), ref error);

                        if (result == 1 && columnHandle != IntPtr.Zero)
                        {
                            try
                            {
                                columnInfo.Add(GetColumnName(columnHandle) ?? $"Column{i}");
                            }
                            finally
                            {
                                _ = Check(libesedb_column_free(ref columnHandle, out error), ref error);
                            }
                        }
                    }
                }
                result = Check(libesedb_table_get_number_of_records(tableHandle, out int numberOfRecords, out error), ref error);
                rows.Add(columnInfo);
                if (result == 1)
                {
                    for (int i = 0; i < numberOfRecords; i++)
                    {
                        result = Check(libesedb_table_get_record(tableHandle, i, out nint recordHandle, out error), ref error);

                        if (result == 1 && recordHandle != IntPtr.Zero)
                        {
                            try
                            {
                                var row = ReadRecord(recordHandle, columnInfo);
                                if (tableName == "SystemIndex_GthrPth")
                                {
                                    List<string> strings = [.. row.Select(o => o?.ToString() ?? string.Empty)];
                                    rows.Add(strings);
                                }
                                else
                                {
                                    rows.Add(row);
                                }
                            }
                            finally
                            {
                                _ = Check(libesedb_record_free(ref recordHandle, out error), ref error);
                            }
                        }
                    }
                }
            }
            finally
            {
                _ = Check(libesedb_table_free(ref tableHandle, out nint error), ref error);
            }
            return rows;
        }

        #region Helper Methods
        private IntPtr GetTableHandle(string tableName)
        {
            int result = Check(libesedb_file_get_number_of_tables(_fileHandle, out int numberOfTables, out nint error), ref error);

            if (result != 1)
                return IntPtr.Zero;
            for (int i = 0; i < numberOfTables; i++)
            {
                result = Check(libesedb_file_get_table(_fileHandle, i, out nint tableHandle, out error), ref error);

                if (result == 1 && tableHandle != IntPtr.Zero)
                {
                    string? currentName = GetTableName(tableHandle);
                    if (currentName == tableName)
                    {
                        return tableHandle;
                    }

                    _ = Check(libesedb_table_free(ref tableHandle, out error), ref error);
                }
            }
            return IntPtr.Zero;
        }

        private static string? GetTableName(IntPtr tableHandle)
        {
            int result = Check(libesedb_table_get_utf8_name_size(tableHandle, out nuint nameSize, out nint error), ref error);
            if (result != 1 || nameSize.ToUInt64() == 0)
                return null;

            byte[] nameBuffer = new byte[nameSize.ToUInt64()];
            result = Check(libesedb_table_get_utf8_name(tableHandle, nameBuffer, nameSize, out error), ref error);
            if (result != 1)
                return null;
            return Encoding.UTF8.GetString(nameBuffer).TrimEnd('\0');
        }

        private static string? GetColumnName(IntPtr columnHandle)
        {
            int result = Check(libesedb_column_get_utf8_name_size(columnHandle, out nuint nameSize, out nint error), ref error);
            if (result != 1 || nameSize.ToUInt64() == 0)
                return null;

            byte[] nameBuffer = new byte[nameSize.ToUInt64()];
            result = Check(libesedb_column_get_utf8_name(columnHandle, nameBuffer, nameSize, out error), ref error);
            if (result != 1)
                return null;
            return Encoding.UTF8.GetString(nameBuffer).TrimEnd('\0');
        }

        private static List<object> ReadRecord(IntPtr recordHandle, List<string> columnNames)
        {
            var row = new List<object>();
            int result = Check(libesedb_record_get_number_of_values(recordHandle, out int numberOfValues, out nint error), ref error);
            if (result != 1)
                return row;

            for (int i = 0; i < numberOfValues && i < columnNames.Count; i++)
            {
                object? value = GetRecordValue(recordHandle, i);
                row.Add(value!);
            }
            return row;
        }

        private static int Check(int result, ref nint error)
        {
            if (error != IntPtr.Zero)
            {
                _ = libesedb_error_free(ref error);
            }
            return result;
        }

        private static object? GetRecordValue(IntPtr recordHandle, int valueIndex)
        {
            int result = Check(libesedb_record_get_column_type(recordHandle, valueIndex, out _, out nint error), ref error);
            if (result != 1)
                return null;

            result = Check(libesedb_record_get_value_utf8_string_size(recordHandle, valueIndex, out nuint valueSize, out error), ref error);
            if (result == 1 && valueSize.ToUInt64() > 0)
            {
                byte[] valueBuffer = new byte[valueSize.ToUInt64()];
                result = Check(libesedb_record_get_value_utf8_string(recordHandle, valueIndex, valueBuffer, valueSize, out error), ref error);
                if (result == 1)
                {
                    return Encoding.UTF8.GetString(valueBuffer).TrimEnd('\0');
                }
            }

            if (GetMultiValueText(recordHandle, valueIndex) is string multiValueText)
            {
                return multiValueText;
            }

            result = Check(libesedb_record_get_value_32bit(recordHandle, valueIndex, out uint value32, out error), ref error);
            if (result == 1)
            {
                return value32;
            }

            result = Check(libesedb_record_get_value_64bit(recordHandle, valueIndex, out ulong value64, out error), ref error);
            if (result == 1)
            {
                return value64;
            }

            result = Check(libesedb_record_get_value_data_size(recordHandle, valueIndex, out valueSize, out error), ref error);
            if (result == 1 && valueSize.ToUInt64() > 0)
            {
                byte[] dataBuffer = new byte[valueSize.ToUInt64()];
                result = Check(libesedb_record_get_value_data(recordHandle, valueIndex, dataBuffer, valueSize, out error), ref error);
                if (result == 1)
                {
                    return dataBuffer;
                }
            }
            return null;
        }

        private static string? GetMultiValueText(IntPtr recordHandle, int valueIndex)
        {
            if (Check(libesedb_record_get_multi_value(recordHandle, valueIndex, out nint multiValue, out nint error), ref error) != 1 || multiValue == IntPtr.Zero)
            {
                return null;
            }
            try
            {
                if (Check(libesedb_multi_value_get_number_of_values(multiValue, out int count, out error), ref error) != 1 || count < 1)
                {
                    return null;
                }
                List<string> values = [];
                for (int i = 0; i < count; i++)
                {
                    if (Check(libesedb_multi_value_get_value_utf8_string_size(multiValue, i, out nuint size, out error), ref error) != 1 || size.ToUInt64() == 0)
                    {
                        return null;
                    }
                    byte[] buffer = new byte[size.ToUInt64()];
                    if (Check(libesedb_multi_value_get_value_utf8_string(multiValue, i, buffer, size, out error), ref error) != 1)
                    {
                        return null;
                    }
                    values.Add(Encoding.UTF8.GetString(buffer).TrimEnd('\0'));
                }
                return string.Join("; ", values);
            }
            finally
            {
                _ = Check(libesedb_multi_value_free(ref multiValue, out error), ref error);
            }
        }

        private void EnsureOpen()
        {
            if (!_isOpen)
                throw new InvalidOperationException("Database is not open. Call Open() first.");
        }

        #endregion

        public void Dispose()
        {
            if (_isOpen && _fileHandle != IntPtr.Zero)
            {
                _ = Check(libesedb_file_close(_fileHandle, out nint error), ref error);
                _ = Check(libesedb_file_free(ref _fileHandle, out error), ref error);
                _isOpen = false;
            }
        }
    }
}
