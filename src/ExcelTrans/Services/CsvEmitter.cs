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
    public static class CsvEmitter
    {
        /// <summary>
        /// Emits the specified w.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="w">The w.</param>
        /// <param name="set">The set.</param>
        public static void Emit<TItem>(TextWriter w, IEnumerable<TItem> set) => Emit(CsvEmitContext.DefaultContext, w, set);
        /// <summary>
        /// Emits the specified context.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="ctx">The context.</param>
        /// <param name="w">The w.</param>
        /// <param name="set">The set.</param>
        public static void Emit<TItem>(CsvEmitContext ctx, TextWriter w, IEnumerable<TItem> set)
        {
            if (ctx == null)
                throw new ArgumentNullException(nameof(ctx));
            if (w == null)
                throw new ArgumentNullException(nameof(w));
            if (set == null)
                throw new ArgumentNullException(nameof(set));
            var hasHeaderRow = (ctx.EmitOptions & CsvEmitOptions.HasHeaderRow) != 0;
            var encodeValues = (ctx.EmitOptions & CsvEmitOptions.EncodeValues) != 0;
            var itemProperties = GetItemProperties<TItem>(hasHeaderRow);

            // header
            var fields = ctx.Fields.Count > 0 ? ctx.Fields : null;
            var b = new StringBuilder();
            if (hasHeaderRow)
            {
                foreach (var itemProperty in itemProperties)
                {
                    // label
                    var displayName = itemProperty.GetDisplayNameAttribute();
                    var name = displayName == null ? itemProperty.Name : displayName.DisplayName;
                    if (fields != null && fields.TryGetValue(itemProperty.Name, out var field) && field != null)
                        if (field.Ignore) continue;
                        else if (field.DisplayName != null) name = field.DisplayName;
                    b.Append(Encode(encodeValues ? EncodeValue(name) : name) + ",");
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
                            string value;
                            var itemValue = itemProperty.GetValue(item);
                            if (fields != null && fields.TryGetValue(itemProperty.Name, out var field) && field != null)
                            {
                                if (field.Ignore) continue;
                                value = field.CustomFieldFormatter == null ? itemValue?.ToString() ?? string.Empty : field.CustomFieldFormatter(field, item, itemValue);
                                if (value.Length == 0)
                                    value = field.DefaultValue ?? string.Empty;
                                if (value.Length != 0)
                                {
                                    var args = field.Args;
                                    if (args != null)
                                    {
                                        if (args.doNotEncode == true)
                                        {
                                            b.Append(value + ",");
                                            continue;
                                        }
                                        if (args.asExcelFunction == true)
                                            value = "=" + value;
                                    }
                                }
                            }
                            else value = itemValue?.ToString() ?? string.Empty;
                            // append value
                            b.Append(Encode(encodeValues ? EncodeValue(value) : value) + ",");
                        }
                        if (b.Length > 0)
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