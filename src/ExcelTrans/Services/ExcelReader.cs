using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelTrans.Services
{
    /// <summary>
    /// Processes the input XLS
    /// </summary>
    public class ExcelReader
    {
        public IEnumerable<T> ExecuteOpenXml<T>(Stream stream, Func<Collection<string>, T> action, int width, int row = 0)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            using (var p = new ExcelPackage(stream))
            {
                var ws = p.Workbook.Worksheets[1];
                ExcelRange range = null;
                while ((range = ws.Cells[row++, width]) != null)
                {
                    var entries = ParseIntoEntries(range);
                    yield return action(entries);
                }
            }
        }

        public IEnumerable<T> ExecuteOpenXml<T>(string path, Func<Collection<string>, T> action, int width, int row = 0)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            var fileInfo = new FileInfo(path);
            using (var p = new ExcelPackage(fileInfo))
            {
                var ws = p.Workbook.Worksheets[1];
                ExcelRange range = null;
                while ((range = ws.Cells[row++, width]) != null)
                {
                    var entries = ParseIntoEntries(range);
                    yield return action(entries);
                }
            }
        }

        public IEnumerable<T> ExecuteBinary<T>(string path, Func<Collection<string>, T> action, int width, int row = 0)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            var sheetName = ExcelFileConnection.GetSheetNames(path).FirstOrDefault();
            using (var table = ExcelFileConnection.GetAllRows(path, 2, false, true, sheetName ?? "Sheet1$"))
                foreach (DataRow r in table.Rows)
                {
                    if (row > 0)
                    {
                        row--;
                        continue;
                    }
                    var entries = ParseIntoEntries(r, width);
                    yield return action(entries);
                }
        }

        Collection<string> ParseIntoEntries(ExcelRange range)
        {
            var list = new Collection<string>();
            foreach (var r in range)
            {
                var str = (r.Value != null ? r.Value.ToString() : null);
                str = str.Trim();
                list.Add(str);
            }
            return list;
        }

        Collection<string> ParseIntoEntries(DataRow r, int width)
        {
            var list = new Collection<string>();
            for (var i = 1; i <= width; i++)
            {
                var str = r.GetExcelValue("F" + i.ToString());
                str = str.Trim();
                list.Add(str);
            }
            return list;
        }
    }
}