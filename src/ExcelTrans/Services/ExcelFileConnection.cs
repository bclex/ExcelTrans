using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ExcelTrans.Services
{
    public static class ExcelFileConnection
    {
        static string GetConnectionString(string file, bool hasHeader, bool allText)
        {
            var extension = Path.GetExtension(file).ToLowerInvariant();
            if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                //uses directory path, not file path
                return string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""text;FMT=Delimited;{1}{2}""",
                    file.Remove(file.IndexOf(Path.GetFileName(file))),
                    hasHeader ? "HDR=YES;" : "HDR=NO;",
                    allText ? "IMEX=1;" : string.Empty);
            else if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                //if this fails, install ACE providers
                return string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;{1}{2}""",
                    file,
                    hasHeader ? "HDR=YES;" : "HDR=NO;",
                    allText ? "IMEX=1;" : string.Empty);
            else
                //assume normal excel
                return string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;{1}{2}""",
                    file,
                    hasHeader ? "HDR=YES;" : "HDR=NO;",
                    allText ? "IMEX=1;" : string.Empty);
        }

        public static string[] GetSheetNames(string file)
        {
            if (string.IsNullOrEmpty(file) || !File.Exists(file))
                return null;
            var connectionString = GetConnectionString(file, false, false);
            using (var connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                var schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new[] { null, null, null, "TABLE" });
                var sheetNames = new List<string>();
                schemaTable.AsEnumerable().ToList().ForEach(x => sheetNames.Add(x.Field<string>("TABLE_NAME")));
                return sheetNames.ToArray();
            }
        }

        public static DataTable GetAllRows(string file, int minColumns, bool hasHeader, bool allText) => GetAllRows(file, minColumns, hasHeader, allText, "Sheet1$");
        public static DataTable GetAllRows(string file, int minColumns, bool hasHeader, bool allText, int rows) => GetAllRows(file, minColumns, hasHeader, allText, rows, "Sheet1$");
        public static DataTable GetAllRows(string file, int minColumns, bool hasHeader, bool allText, string sheetName) => ExecuteSqlAgainstFile(file, minColumns, hasHeader, allText, $"Select * From [{sheetName}]");
        public static DataTable GetAllRows(string file, int minColumns, bool hasHeader, bool allText, string sheetName, string sql) => ExecuteSqlAgainstFile(file, minColumns, hasHeader, allText, string.Format(sql, sheetName));
        public static DataTable GetAllRows(string file, int minColumns, bool hasHeader, bool allText, int rows, string sheetName) => ExecuteSqlAgainstFile(file, minColumns, hasHeader, allText, $"Select Top {rows} * From [{sheetName}]");

        static DataTable ExecuteSqlAgainstFile(string file, int minColumns, bool hasHeader, bool allText, string sql)
        {
            if (string.IsNullOrEmpty(file) || !File.Exists(file))
                return null;
            var connectionString = GetConnectionString(file, hasHeader, allText);
            var table = new DataTable();
            using (var connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (var adapter = new OleDbDataAdapter(sql, connection))
                    adapter.Fill(table);
            }
            if (table.Columns.Count < minColumns)
                throw new ArgumentOutOfRangeException(nameof(minColumns));
            return table;
        }

        public static string GetExcelValue(this DataRow row, params string[] fieldNames)
        {
            var columns = row.Table.Columns;
            var fieldName = fieldNames.Where(x => columns.Contains(x) && !row.IsNull(columns[x])).FirstOrDefault();
            if (string.IsNullOrEmpty(fieldName))
                return string.Empty;
            var column = row.Table.Columns[fieldName];
            switch (column.DataType.ToString())
            {
                case "System.DateTime": return row[fieldName] == DBNull.Value ? string.Empty : ((DateTime)row[fieldName]).ToString("MM/dd/yyyy");
                case "System.String": return row.Field<string>(fieldName) ?? string.Empty;
                default: return row[fieldName] == DBNull.Value || row[fieldName] == null ? string.Empty : row[fieldName].ToString();
            }
        }
    }
}