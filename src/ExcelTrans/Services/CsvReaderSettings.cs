using System;
using System.Globalization;

namespace ExcelTrans.Services
{
    /// <summary>
    /// Class CsvReaderSettings.
    /// </summary>
    public class CsvReaderSettings
    {
        string _delimiterAsString;

        /// <summary>
        /// Gets or sets the delimiter as string.
        /// </summary>
        /// <value>
        /// The delimiter as string.
        /// </value>
        public string DelimiterAsString
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
    }
}