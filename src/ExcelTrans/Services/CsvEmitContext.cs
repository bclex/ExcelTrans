using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelTrans.Services
{
    /// <summary>
    /// CsvEmitContext
    /// </summary>
    public class CsvEmitContext
    {
        /// <summary>
        /// CsvEmitFilterMode
        /// </summary>
        public enum CsvEmitFilterMode
        {
            /// <summary>
            /// ExceptionsInFields
            /// </summary>
            ExceptionsInFields,
            /// <summary>
            /// InclusionsInFields
            /// </summary>
            InclusionsInFields,
        }

        /// <summary>
        /// CsvEmitOptions
        /// </summary>
        public enum CsvEmitOptions
        {
            /// <summary>
            /// HasHeaderRow
            /// </summary>
            HasHeaderRow = 0x1,
            /// <summary>
            /// IncludeFields
            /// </summary>
            IncludeFields = 0x2,
            /// <summary>
            /// EncodeValues
            /// </summary>
            EncodeValues = 0x4,
        }

        /// <summary>
        /// CsvEmitFieldCollection
        /// </summary>
        public class CsvEmitFieldCollection : KeyedCollection<string, CsvEmitField>
        {
            /// <summary>
            /// Tries the get value.
            /// </summary>
            /// <param name="name">The name.</param>
            /// <param name="field">The field.</param>
            /// <returns></returns>
            public bool TryGetValue(string name, out CsvEmitField field)
            {
                if (!Contains(name))
                {
                    field = null;
                    return false;
                }
                field = base[name];
                return true;
            }

            /// <summary>
            /// When implemented in a derived class, extracts the key from the specified element.
            /// </summary>
            /// <param name="item">The element from which to extract the key.</param>
            /// <returns>
            /// The key for the specified element.
            /// </returns>
            protected override string GetKeyForItem(CsvEmitField item) { return item.Name; }
        }

        /// <summary>
        /// DefaultContext
        /// </summary>
        public static readonly CsvEmitContext DefaultContext = new CsvEmitContext { };

        /// <summary>
        /// Initializes a new instance of the <see cref="CsvEmitContext"/> class.
        /// </summary>
        public CsvEmitContext()
        {
            EmitOptions = CsvEmitOptions.IncludeFields | CsvEmitOptions.HasHeaderRow | CsvEmitOptions.EncodeValues;
            FilterMode = CsvEmitFilterMode.ExceptionsInFields;
            Fields = new CsvEmitFieldCollection();
            FlushAt = 500;
        }

        /// <summary>
        /// Gets or sets the <see cref="System.Boolean"/> with the specified option.
        /// </summary>
        public bool this[CsvEmitOptions option]
        {
            get { return ((EmitOptions & option) == option); }
            set { EmitOptions = (value ? EmitOptions | option : EmitOptions & ~option); }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance has header row.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance has header row; otherwise, <c>false</c>.
        /// </value>
        public bool HasHeaderRow
        {
            get { return this[CsvEmitOptions.HasHeaderRow]; }
            set { this[CsvEmitOptions.HasHeaderRow] = value; }
        }

        /// <summary>
        /// Gets or sets the filter mode.
        /// </summary>
        /// <value>
        /// The filter mode.
        /// </value>
        public CsvEmitFilterMode FilterMode { get; set; }
        /// <summary>
        /// Gets the fields.
        /// </summary>
        public CsvEmitFieldCollection Fields { get; private set; }
        /// <summary>
        /// Gets or sets the emit options.
        /// </summary>
        /// <value>
        /// The emit options.
        /// </value>
        public CsvEmitOptions EmitOptions { get; set; }
        /// <summary>
        /// Gets or sets the flush at.
        /// </summary>
        /// <value>
        /// The flush at.
        /// </value>
        public int FlushAt { get; set; }
        /// <summary>
        /// Gets or sets the on flush.
        /// </summary>
        /// <value>
        /// The on flush.
        /// </value>
        public Func<IEnumerable<object>, IEnumerable<object>> BeforeFlush { get; set; }
    }
}