using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace ExcelTrans.Services
{
    /// <summary>
    /// Processes the input XLS
    /// </summary>
    public static class ExcelReader
    {
        public static IEnumerable<T> ExecuteOpenXml<T>(Stream stream, Func<Collection<string>, T> action, int width, int row = 0)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            using (var p = new ExcelPackage(stream))
            {
                var ws = p.Workbook.Worksheets[1];
                ExcelRange range = null;
                var list = new Collection<string>();
                while ((range = ws.Cells[row++, width]) != null)
                {
                    var entries = ParseIntoEntries(list, range);
                    yield return action(entries);
                }
            }
        }

        public static IEnumerable<T> ExecuteOpenXml<T>(string path, Func<Collection<string>, T> action, int width, int row = 0)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            var fileInfo = new FileInfo(path);
            using (var p = new ExcelPackage(fileInfo))
            {
                var ws = p.Workbook.Worksheets[1];
                ExcelRange range = null;
                var list = new Collection<string>();
                while ((range = ws.Cells[row++, width]) != null)
                {
                    var entries = ParseIntoEntries(list, range);
                    yield return action(entries);
                }
            }
        }

        public static IEnumerable<T> ExecuteRawXml<T>(Stream stream, Func<Collection<string>, T> action, int width, int row = 0)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            var xml_ = new StreamReader(stream).ReadToEnd();
            int idx = 0, idx2;
            while (true)
            {
                idx = xml_.IndexOf("<table", idx);
                idx2 = idx != -1 ? xml_.IndexOf("</table>", idx) : -1;
                if (idx2 == -1)
                    break;
                var xml = xml_.Substring(idx, idx2 - idx + 8).Replace(" nowrap", "").Replace("&", "&amp;");
                XDocument doc;
                try
                {
                    using (var s = new StringReader(xml))
                        doc = XDocument.Load(s);
                }
                catch (Exception e) { throw ParsingException(e, xml); }
                var list = new Collection<string>();
                foreach (var r in doc.Descendants("tr"))
                {
                    if (row > 0)
                    {
                        row--;
                        continue;
                    }
                    var entries = ParseIntoEntries(list, r, width);
                    yield return action(entries);
                }
                // next
                idx = idx2;
            }
        }

        public static IEnumerable<T> ExecuteRaw2Xml<T>(Stream stream, Func<Collection<string>, T> action, int width, int row = 0)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            var xml_ = new StreamReader(stream).ReadToEnd();
            int idx = 0, idx2;
            while (true)
            {
                idx = xml_.IndexOf("<table", idx);
                idx2 = idx != -1 ? xml_.IndexOf("</table>", idx) : -1;
                if (idx2 == -1)
                    break;
                var xml = xml_.Substring(idx, idx2 - idx + 8).Replace(" nowrap", "").Replace("&", "&amp;");
                var list = new Collection<string>();
                using (var r = XmlReader.Create(new StringReader(xml)))
                {
                    r.MoveToContent();
                    while (r.Read())
                    {
                        if (r.NodeType != XmlNodeType.Element || r.Name != "tr")
                            continue;
                        if (row > 0)
                        {
                            row--;
                            continue;
                        }
                        var entries = ParseIntoEntries(list, r, width);
                        yield return action(entries);
                    }
                }
                // next
                idx = idx2;
            }
        }

        static Exception ParsingException(Exception e, string xml)
        {
            var msg = e.Message;
            if (!msg.Contains("Line") || !msg.Contains("position"))
                return e;
            var idx = msg.IndexOf("Line"); var idx2 = msg.IndexOf(",", idx); var line = int.Parse(msg.Substring(idx + 4, idx2 - idx - 4));
            if (line != 1)
                return e;
            idx = msg.IndexOf("position"); idx2 = msg.IndexOf(".", idx); var position = int.Parse(msg.Substring(idx + 8, idx2 - idx - 8));
            var error = xml.Substring(position - 30, 30) + "!!" + xml.Substring(position, 20);
            return new ArgumentOutOfRangeException(msg, error);
        }

        public static IEnumerable<T> ExecuteRawXml<T>(string path, Func<Collection<string>, T> action, int width, int row = 0) => ExecuteRawXml(File.OpenRead(path), action, width, row);

        public static IEnumerable<T> ExecuteBinary<T>(string path, Func<Collection<string>, T> action, int width, int row = 0)
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
                    var list = new Collection<string>();
                    var entries = ParseIntoEntries(list, r, width);
                    yield return action(entries);
                }
        }

        static Collection<string> ParseIntoEntries(Collection<string> list, XmlReader r, int width)
        {
            list.Clear();
            while (r.Read())
            {
                if (r.NodeType == XmlNodeType.EndElement && r.Name == "tr")
                {
                    for (var i = list.Count; i < width; i++)
                        list.Add(null);
                    return list;
                }
                if (r.NodeType != XmlNodeType.Text)
                    continue;
                list.Add(r.Value.Trim());
            }
            throw new InvalidOperationException();
        }

        static Collection<string> ParseIntoEntries(Collection<string> list, XElement r, int width)
        {
            var cols = r.Descendants("th").Concat(r.Descendants("td")).ToArray();
            list.Clear();
            for (var i = 0; i < width; i++)
            {
                if (i >= cols.Length)
                {
                    list.Add(null);
                    continue;
                }
                var col = cols[i];
                list.Add(col.Value.Trim());
            }
            return list;
        }

        static Collection<string> ParseIntoEntries(Collection<string> list, ExcelRange range)
        {
            list.Clear();
            foreach (var r in range)
            {
                var str = r.Value?.ToString();
                list.Add(str.Trim());
            }
            return list;
        }

        static Collection<string> ParseIntoEntries(Collection<string> list, DataRow r, int width)
        {
            list.Clear();
            for (var i = 1; i <= width; i++)
            {
                var str = r.GetExcelValue("F" + i.ToString());
                list.Add(str.Trim());
            }
            return list;
        }
    }
}