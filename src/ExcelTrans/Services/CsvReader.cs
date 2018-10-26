using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelTrans.Services
{
    /// <summary>
    /// Processes the input CSV
    /// </summary>
    public class CsvReader
    {
        /// <summary>
        /// The delimiter as string
        /// </summary>
        string _delimiterAsString;

        /// <summary>
        /// Gets or sets the delimiter as string.
        /// </summary>
        /// <value>
        /// The delimiter as string.
        /// </value>
        string DelimiterAsString
        {
            get
            {
                if (_delimiterAsString == null)
                    try
                    {
                        _delimiterAsString = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
                        if (_delimiterAsString.Length != 1)
                            _delimiterAsString = ",";
                    }
                    catch (Exception) { _delimiterAsString = ","; }
                return _delimiterAsString;
            }
            set { _delimiterAsString = value; }
        }

        /// <summary>
        /// Executes the specified reader.
        /// </summary>
        /// <param name="reader">The reader instance.</param>
        /// <param name="action">The logic to execute.</param>
        /// <exception cref="System.ArgumentNullException">If the reader instance is null</exception>
        public IEnumerable<T> Execute<T>(TextReader reader, Func<Collection<string>, T> action)
        {
            if (reader == null)
                throw new ArgumentNullException(nameof(reader));
            var delimiter = DelimiterAsString[0];
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                var entries = !string.IsNullOrEmpty(line.Trim()) ? ParseLineIntoEntries(delimiter, line, () => reader.ReadLine()) : null;
                yield return action(entries);
            }
        }

        /// <summary>
        /// Parses the line into entries.
        /// </summary>
        /// <param name="line">The line to parse.</param>
        /// <returns>A collection of columns</returns>
        Collection<string> ParseLineIntoEntries(char delimiter, string line, Func<string> readLine)
        {
            var list = new Collection<string>();
            var lineArray = line.ToCharArray();
            var inQuote = false;
            var b = new StringBuilder();
            for (var i = 0; i < line.Length; i++)
            {
                if (!inQuote && b.Length == 0)
                {
                    if (char.IsWhiteSpace(lineArray[i])) continue;
                    if (lineArray[i] == '"') { inQuote = true; continue; }
                }
                if (inQuote && lineArray[i] == '"')
                {
                    if (i + 1 < line.Length && lineArray[i + 1] == '"') i++; // double quote error
                    else { if (i + 1 < line.Length && lineArray[i + 1] != delimiter) return null; inQuote = false; continue; } // broken quote error
                }
                if (inQuote || lineArray[i] != delimiter)
                {
                    b.Append(lineArray[i]);
                    if (inQuote && i + 1 == line.Length) { b.Append("\r\n"); i = -1; line = readLine(); lineArray = line.ToCharArray(); } // line spill
                }
                else { list.Add(b.ToString().Trim()); b.Length = 0; }
            }
            list.Add(b.ToString().Trim());
            return list;
        }
    }
}