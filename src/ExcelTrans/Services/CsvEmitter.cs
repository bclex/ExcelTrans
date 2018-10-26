using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using CsvEmitOptions = ExcelTrans.Services.CsvEmitContext.CsvEmitOptions;

namespace ExcelTrans.Services
{
    /// <summary>
    /// CsvEmitter
    /// </summary>
    public class CsvEmitter
    {
        /// <summary>
        /// Emits the specified w.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="w">The w.</param>
        /// <param name="set">The set.</param>
        public void Emit<TItem>(TextWriter w, IEnumerable<TItem> set) { Emit<TItem>(CsvEmitContext.DefaultContext, w, set); }
        /// <summary>
        /// Emits the specified context.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="ctx">The context.</param>
        /// <param name="w">The w.</param>
        /// <param name="set">The set.</param>
        public void Emit<TItem>(CsvEmitContext ctx, TextWriter w, IEnumerable<TItem> set)
        {
            if (ctx == null)
                throw new ArgumentNullException(nameof(ctx));
            if (w == null)
                throw new ArgumentNullException(nameof(w));
            if (set == null)
                throw new ArgumentNullException(nameof(set));
            var itemProperties = GetItemProperties<TItem>((ctx.EmitOptions & CsvEmitOptions.HasHeaderRow) != 0);
            var shouldEncodeValues = (ctx.EmitOptions & CsvEmitOptions.EncodeValues) != 0;

            // header
            var fields = ctx.Fields.Count > 0 ? ctx.Fields : null;
            var b = new StringBuilder();
            if ((ctx.EmitOptions & CsvEmitOptions.HasHeaderRow) != 0)
            {
                foreach (var itemProperty in itemProperties)
                {
                    // label
                    var displayName = itemProperty.GetDisplayNameAttribute(); // GetCustomAttribute<DisplayNameAttribute>();
                    var name = displayName == null ? itemProperty.Name : displayName.DisplayName;
                    if (fields != null && fields.TryGetValue(itemProperty.Name, out CsvEmitField field) && field != null)
                        if (field.Ignore) continue;
                        else if (field.DisplayName != null) name = field.DisplayName;
                    b.Append(Encode(shouldEncodeValues ? EncodeValue(name) : name) + ",");
                }
                if (b.Length > 0)
                    b.Length--;
                w.Write(b.ToString() + Environment.NewLine);
            }

            // rows
            try
            {
                foreach (var group in set.Cast<object>().GroupAt(ctx.FlushAt))
                {
                    var newGroup = ctx.BeforeFlush == null ? group : ctx.BeforeFlush(group);
                    if (newGroup == null)
                        return;
                    foreach (var item in newGroup)
                    {
                        b.Length = 0;
                        foreach (var itemProperty in itemProperties)
                        {
                            // value
                            string valueAsText;
                            var value = itemProperty.GetValue(item);
                            if (fields != null && fields.TryGetValue(itemProperty.Name, out CsvEmitField field) && field != null)
                            {
                                if (field.Ignore)
                                    continue;
                                var fieldFormatter = field.CustomFieldFormatter;
                                if (fieldFormatter == null)
                                {
                                    // default formatter
                                    valueAsText = value != null ? value.ToString() : string.Empty;
                                    if (valueAsText.Length == 0)
                                        valueAsText = field.DefaultValue;
                                    continue;
                                }
                                // formatter
                                valueAsText = fieldFormatter(field, item, value);
                                if (!string.IsNullOrEmpty(valueAsText))
                                {
                                    var args = field.Args;
                                    if (args != null)
                                    {
                                        if (args.doNotEncode == true)
                                        {
                                            b.Append(valueAsText + ",");
                                            continue;
                                        }
                                        if (args.asExcelFunction == true)
                                            valueAsText = "=" + valueAsText;
                                    }
                                }
                            }
                            else valueAsText = (value != null ? value.ToString() : string.Empty);
                            // append value
                            b.Append(Encode(shouldEncodeValues ? EncodeValue(valueAsText) : valueAsText) + ",");
                        }
                        b.Length--;
                        w.Write(b.ToString() + Environment.NewLine);
                    }
                    w.Flush();
                }
            }
            finally { w.Flush(); }
        }

        class ItemInfo
        {
            public string Name;
            public Func<object, object> GetValue;
            public Func<DisplayNameAttribute> GetDisplayNameAttribute;
        }

        static List<ItemInfo> GetItemProperties<T>(bool includeFields)
        {
            var items = typeof(T).GetProperties().Select(x => new ItemInfo
            {
                Name = x.Name,
                GetValue = x.GetValue,
                GetDisplayNameAttribute = () => x.GetCustomAttribute<DisplayNameAttribute>(),
            }).ToList();
            if (includeFields)
                items.AddRange(typeof(T).GetFields().Select(x => new ItemInfo
                {
                    Name = x.Name,
                    GetValue = x.GetValue,
                    GetDisplayNameAttribute = () => x.GetCustomAttribute<DisplayNameAttribute>(),
                }));
            return items;
        }
        static string Encode(string value) => value;
        static string EncodeValue(string value) => string.IsNullOrEmpty(value) ? "\"\"" : "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}