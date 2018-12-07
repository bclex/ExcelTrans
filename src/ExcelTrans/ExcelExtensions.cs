using ExcelTrans.Commands;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace ExcelTrans
{
    public static class ExcelExtensions
    {
        public static object Execute(this IExcelContext ctx, IExcelCommand[] cmds, out Action after)
        {
            var si = ctx.GetCtx();
            var afterCmds = new List<IExcelCommand>();
            foreach (var cmd in cmds)
                if (cmd.When <= When.Before) cmd.Execute(ctx);
                else afterCmds.Add(cmd);
            after = afterCmds.Count > 0 ? () => { foreach (var cmd in afterCmds) cmd.Execute(ctx); } : (Action)null;
            return si;
        }

        public static CommandRtn ExecuteRow(this IExcelContext ctx, When when, Collection<string> s, out Action after)
        {
            var cr = CommandRtn.None;
            var afterActions = new List<Action>();
            foreach (var cmd in ctx.Cmds.SelectMany(z => z.Item1.Where(x => (x.When & when) == when)))
            {
                var r = cmd.Func(ctx, s);
                if ((r & CommandRtn.Execute) == CommandRtn.Execute)
                {
                    ctx.SetCtx(ctx.Execute(cmd.Cmds, out Action subAfter));
                    if (subAfter != null) afterActions.Add(subAfter);
                }
                cr |= r;
            }
            after = afterActions.Count > 0 ? () => { foreach (var action in afterActions) action(); } : (Action)null;
            return cr;
        }

        public static CommandRtn ExecuteCol(this IExcelContext ctx, Collection<string> s, object v, int i, out Action after)
        {
            var cr = CommandRtn.None;
            var afterActions = new List<Action>();
            foreach (var cmd in ctx.Cmds.SelectMany(z => z.Item2))
            {
                var r = cmd.Func(ctx, s, v, i);
                if ((r & CommandRtn.Execute) == CommandRtn.Execute)
                {
                    ctx.SetCtx(ctx.Execute(cmd.Cmds, out Action subAfter));
                    if (subAfter != null) afterActions.Add(subAfter);
                }
                cr |= r;
            }
            after = afterActions.Count > 0 ? () => { foreach (var action in afterActions) action(); } : (Action)null;
            return cr;
        }

        public static void WriteRowFirstSet(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.FirstSet, s, out Action after);
        public static void WriteRowFirst(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.First, s, out Action after);

        public static void WriteRow(this IExcelContext ctx, Collection<string> s)
        {
            var ws = ((ExcelContext)ctx).EnsureWorksheet();
            // execute-row-before
            var cr = ctx.ExecuteRow(When.Before, s, out Action after);
            if ((cr & CommandRtn.Skip) == CommandRtn.Skip)
                return;
            ctx.X = ctx.XStart;
            for (var i = 0; i < s.Count; i++)
            {
                ctx.CsvX = i + 1;
                var v = s[i].ParseValue();
                // execute-col
                cr = ctx.ExecuteCol(s, v, i, out Action subAfter);
                if ((cr & CommandRtn.Skip) == CommandRtn.Skip)
                    continue;
                if ((cr & CommandRtn.Formula) != CommandRtn.Formula) ws.SetValue(ctx.Y, ctx.X, v);
                else ws.Cells[ctx.Y, ctx.X].Formula = s[i];
                //if (v is DateTime) ws.Cells[ExcelCellBase.GetAddress(ctx.Y, ctx.X)].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                subAfter?.Invoke();
                ctx.X += ctx.DeltaX;
            }
            after?.Invoke();
            ctx.Y += ctx.DeltaY;
            // execute-row-after
            ctx.ExecuteRow(When.After, s, out Action after2);
        }

        public static void WriteRowLast(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.Last, s, out Action after);
        public static void WriteRowLastSet(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.LastSet, s, out Action after);

        static Color ToColor(string name)
        {
            var propMethod = typeof(Color).GetProperty(name, BindingFlags.Public | BindingFlags.Static);
            if (propMethod == null)
                throw new ArgumentNullException(nameof(name), $"Unable to find color {name}");
            return (Color)propMethod.GetValue(null);
        }

        // https://stackoverflow.com/questions/40209636/epplus-number-format/40214134
        public static void CellsStyle(this IExcelContext ctx, int row, int col, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(row, col), styles);
        public static void CellsStyle(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(r, 0, 0), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, int row, int col, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(r, row, col), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), styles);
        public static void CellsStyle(this IExcelContext ctx, string cells, string[] styles)
        {
            string NumberformatPrec(string prec, string defaultPrec) => string.IsNullOrEmpty(prec) ? defaultPrec : $"0.{new String('0', int.Parse(prec))}";
            var range = ((ExcelContext)ctx).WS.Cells[ExcelService.DecodeAddress(ctx, cells)];
            foreach (var style in styles)
            {
                // number-format
                if (style.StartsWith("n"))
                {
                    // https://support.office.com/en-us/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
                    if (style.StartsWith("n:")) range.Style.Numberformat.Format = style.Substring(2);
                    else if (style.StartsWith("n$")) range.Style.Numberformat.Format = $"_(\"$\"* #,##{NumberformatPrec(style.Substring(2), "0.00")}_);_(\"$\"* \\(#,##{NumberformatPrec(style.Substring(2), "0.00")}\\);_(\"$\"* \" - \"??_);_(@_)"; // "_-$* #,##{NumberformatPrec(style.Substring(2), "0.00")}_-;-$* #,##{NumberformatPrec(style.Substring(2), "0.00")}_-;_-$* \"-\"??_-;_-@_-";
                    else if (style.StartsWith("n%")) range.Style.Numberformat.Format = $"{NumberformatPrec(style.Substring(2), "0")}%";
                    else if (style.StartsWith("n,")) range.Style.Numberformat.Format = $"_(* #,##{NumberformatPrec(style.Substring(2), "0.00")}_);_(* \\(#,##{NumberformatPrec(style.Substring(2), "0.00")}\\);_(* \"-\"??_);_(@_)";
                    else if (style == "nd") range.Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    else throw new InvalidOperationException($"{style} not defined");
                }
                // font
                else if (style.StartsWith("f"))
                {
                    if (style.StartsWith("f:")) range.Style.Font.Name = style.Substring(2);
                    else if (style.StartsWith("fx")) range.Style.Font.Size = float.Parse(style.Substring(2));
                    else if (style.StartsWith("ff")) range.Style.Font.Family = int.Parse(style.Substring(2));
                    else if (style.StartsWith("fc:")) range.Style.Font.Color.SetColor(ToColor(style.Substring(3)));
                    else if (style.StartsWith("fs:")) range.Style.Font.Scheme = style.Substring(2);
                    else if (style == "fB") range.Style.Font.Bold = true;
                    else if (style == "fb") range.Style.Font.Bold = false;
                    else if (style == "fI") range.Style.Font.Italic = true;
                    else if (style == "fi") range.Style.Font.Italic = false;
                    else if (style == "fS") range.Style.Font.Strike = true;
                    else if (style == "fs") range.Style.Font.Strike = false;
                    else if (style == "f_") range.Style.Font.UnderLine = true;
                    else if (style == "f!_") range.Style.Font.UnderLine = false;
                    //else if (style == "") range.Style.Font.UnderLineType = ?;
                    else if (style.StartsWith("fv")) range.Style.Font.VerticalAlign = (ExcelVerticalAlignmentFont)int.Parse(style.Substring(2));
                    else throw new InvalidOperationException($"{style} not defined");
                }
                // fill
                else if (style.StartsWith("l"))
                {
                    if (style.StartsWith("lc:"))
                    {
                        if (range.Style.Fill.PatternType == ExcelFillStyle.None) range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(ToColor(style.Substring(3)));
                    }
                    else if (style.StartsWith("lf")) range.Style.Fill.PatternType = (ExcelFillStyle)int.Parse(style.Substring(2));
                }
                // border
                else if (style.StartsWith("b"))
                {
                    if (style.StartsWith("bl")) range.Style.Border.Left.Style = (ExcelBorderStyle)int.Parse(style.Substring(2));
                    else if (style.StartsWith("br")) range.Style.Border.Right.Style = (ExcelBorderStyle)int.Parse(style.Substring(2));
                    else if (style.StartsWith("bt")) range.Style.Border.Top.Style = (ExcelBorderStyle)int.Parse(style.Substring(2));
                    else if (style.StartsWith("bb")) range.Style.Border.Bottom.Style = (ExcelBorderStyle)int.Parse(style.Substring(2));
                    else if (style.StartsWith("bd")) range.Style.Border.Diagonal.Style = (ExcelBorderStyle)int.Parse(style.Substring(2));
                    else if (style == "bdU") range.Style.Border.DiagonalUp = true;
                    else if (style == "bdu") range.Style.Border.DiagonalUp = false;
                    else if (style == "bdD") range.Style.Border.DiagonalDown = true;
                    else if (style == "bdd") range.Style.Border.DiagonalDown = false;
                    else if (style.StartsWith("ba")) range.Style.Border.BorderAround((ExcelBorderStyle)int.Parse(style.Substring(2))); // add color option
                    else throw new InvalidOperationException($"{style} not defined");
                }
                // horizontal-alignment
                else if (style.StartsWith("ha"))
                {
                    range.Style.HorizontalAlignment = (ExcelHorizontalAlignment)int.Parse(style.Substring(2));
                }
                // vertical-alignment
                else if (style.StartsWith("va"))
                {
                    range.Style.VerticalAlignment = (ExcelVerticalAlignment)int.Parse(style.Substring(2));
                }
                // vertical-alignment
                else if (style.StartsWith("W")) range.Style.WrapText = true;
                else if (style.StartsWith("w")) range.Style.WrapText = false;
                else throw new InvalidOperationException($"{style} not defined");
            }
        }

        public static void CellsValue(this IExcelContext ctx, int row, int col, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(row, col), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, Address r, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(r, 0, 0), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, Address r, int row, int col, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(r, row, col), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, string cells, object value, CellValueKind valueKind = CellValueKind.Value)
        {
            var range = ((ExcelContext)ctx).WS.Cells[ExcelService.DecodeAddress(ctx, cells)];
            switch (valueKind)
            {
                case CellValueKind.Value: range.Value = value; break;
                case CellValueKind.AutoFilter: range.AutoFilter = value.CastValue<bool>(); break;
                case CellValueKind.AutoFitColumns: range.AutoFitColumns(); break;
                case CellValueKind.Comment: range.Comment.Text = (string)value; break;
                case CellValueKind.CommentMore: break;
                case CellValueKind.ConditionalFormattingMore: break;
                case CellValueKind.Copy:
                    var range2 = ((ExcelContext)ctx).WS.Cells[ExcelService.DecodeAddress(ctx, (string)value)];
                    range.Copy(range2); break;
                case CellValueKind.Formula: range.Formula = (string)value; break;
                case CellValueKind.FormulaR1C1: range.FormulaR1C1 = (string)value; break;
                case CellValueKind.Hyperlink: range.Hyperlink = new Uri((string)value); break;
                case CellValueKind.Merge: range.Merge = value.CastValue<bool>(); break;
                case CellValueKind.RichText: range.RichText.Add((string)value); break;
                case CellValueKind.RichTextClear: range.RichText.Clear(); break;
                case CellValueKind.StyleName: range.StyleName = (string)value; break;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
            if (value is DateTime) range.Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
        }

        public static object GetCellsValue(this IExcelContext ctx, int row, int col, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(row, col), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, Address r, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(r, 0, 0), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, Address r, int row, int col, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(r, row, col), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, string cells, CellValueKind valueKind = CellValueKind.Value)
        {
            var range = ((ExcelContext)ctx).WS.Cells[ExcelService.DecodeAddress(ctx, cells)];
            switch (valueKind)
            {
                case CellValueKind.Value: return range.Value;
                case CellValueKind.Text: return range.Text;
                case CellValueKind.AutoFilter: return range.AutoFilter;
                case CellValueKind.Comment: return range.Comment.Text;
                case CellValueKind.ConditionalFormattingMore: return null;
                case CellValueKind.Formula: return range.Formula;
                case CellValueKind.FormulaR1C1: return range.FormulaR1C1;
                case CellValueKind.Hyperlink: return range.Hyperlink;
                case CellValueKind.Merge: return range.Merge;
                case CellValueKind.StyleName: return range.StyleName;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        #region Column

        public static void DeleteColumn(this IExcelContext ctx, int column) => ((ExcelContext)ctx).WS.DeleteColumn(column);
        public static void DeleteColumn(this IExcelContext ctx, int columnFrom, int columns) => ((ExcelContext)ctx).WS.DeleteColumn(columnFrom, columns);

        public static void InsertColumn(this IExcelContext ctx, int columnFrom, int columns) => ((ExcelContext)ctx).WS.InsertColumn(columnFrom, columns);
        public static void InsertColumn(this IExcelContext ctx, int columnFrom, int columns, int copyStylesFromColumn) => ((ExcelContext)ctx).WS.InsertColumn(columnFrom, columns, copyStylesFromColumn);

        public static void ColumnValue(this IExcelContext ctx, string col, object value, ColumnValueKind valueKind) => ColumnValue(ctx, ExcelService.ColToInt(col), value, valueKind);
        public static void ColumnValue(this IExcelContext ctx, int col, object value, ColumnValueKind valueKind)
        {
            var column = ((ExcelContext)ctx).WS.Column(col);
            switch (valueKind)
            {
                case ColumnValueKind.AutoFit: column.AutoFit(); break;
                case ColumnValueKind.BestFit: column.BestFit = value.CastValue<bool>(); break;
                case ColumnValueKind.Merged: column.Merged = value.CastValue<bool>(); break;
                case ColumnValueKind.Width: column.Width = value.CastValue<double>(); break;
                case ColumnValueKind.TrueWidth: column.SetTrueColumnWidth(value.CastValue<double>()); break;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        public static object GetColumnValue(this IExcelContext ctx, string col, ColumnValueKind valueKind) => GetColumnValue(ctx, ExcelService.ColToInt(col), valueKind);
        public static object GetColumnValue(this IExcelContext ctx, int col, ColumnValueKind valueKind)
        {
            var column = ((ExcelContext)ctx).WS.Column(col);
            switch (valueKind)
            {
                case ColumnValueKind.BestFit: return column.BestFit;
                case ColumnValueKind.Merged: return column.Merged;
                case ColumnValueKind.Width: return column.Width;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        public static void SetTrueColumnWidth(this ExcelColumn column, double width)
        {
            // Deduce what the column width would really get set to.
            var z = width >= (1 + 2 / 3)
                ? Math.Round((Math.Round(7 * (width - 1 / 256), 0) - 5) / 7, 2)
                : Math.Round((Math.Round(12 * (width - 1 / 256), 0) - Math.Round(5 * width, 0)) / 12, 2);

            // How far off? (will be less than 1)
            var errorAmt = width - z;

            // Calculate what amount to tack onto the original amount to result in the closest possible setting.
            var adj = width >= 1 + 2 / 3
                ? Math.Round(7 * errorAmt - 7 / 256, 0) / 7
                : Math.Round(12 * errorAmt - 12 / 256, 0) / 12 + (2 / 12);

            // Set width to a scaled-value that should result in the nearest possible value to the true desired setting.
            if (z > 0)
            {
                column.Width = width + adj;
                return;
            }
            column.Width = 0d;
        }

        #endregion

        #region Row

        public static void DeleteRow(this IExcelContext ctx, int row) => ((ExcelContext)ctx).WS.DeleteRow(row);
        public static void DeleteRow(this IExcelContext ctx, int rowFrom, int rows) => ((ExcelContext)ctx).WS.DeleteRow(rowFrom, rows);

        public static void InsertRow(this IExcelContext ctx, int rowFrom, int rows) => ((ExcelContext)ctx).WS.InsertRow(rowFrom, rows);
        public static void InsertRow(this IExcelContext ctx, int rowFrom, int rows, int copyStylesFromRow) => ((ExcelContext)ctx).WS.InsertRow(rowFrom, rows, copyStylesFromRow);

        public static void RowValue(this IExcelContext ctx, string row, object value, RowValueKind valueKind) => RowValue(ctx, ExcelService.RowToInt(row), value, valueKind);
        public static void RowValue(this IExcelContext ctx, int row, object value, RowValueKind valueKind)
        {
            var row_ = ((ExcelContext)ctx).WS.Row(row);
            switch (valueKind)
            {
                case RowValueKind.Collapsed: row_.Collapsed = value.CastValue<bool>(); break;
                case RowValueKind.CustomHeight: row_.CustomHeight = value.CastValue<bool>(); break;
                case RowValueKind.Height: row_.Height = value.CastValue<double>(); break;
                case RowValueKind.Hidden: row_.Hidden = value.CastValue<bool>(); break;
                case RowValueKind.Merged: row_.Merged = value.CastValue<bool>(); break;
                case RowValueKind.OutlineLevel: row_.OutlineLevel = value.CastValue<int>(); break;
                case RowValueKind.PageBreak: row_.PageBreak = value.CastValue<bool>(); break;
                case RowValueKind.Phonetic: row_.Phonetic = value.CastValue<bool>(); break;
                case RowValueKind.StyleName: row_.StyleName = value.CastValue<string>(); break;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        public static object GetRowValue(this IExcelContext ctx, string row, RowValueKind valueKind) => GetRowValue(ctx, ExcelService.RowToInt(row), valueKind);
        public static object GetRowValue(this IExcelContext ctx, int row, RowValueKind valueKind)
        {
            var row_ = ((ExcelContext)ctx).WS.Row(row);
            switch (valueKind)
            {
                case RowValueKind.Collapsed: return row_.Collapsed;
                case RowValueKind.CustomHeight: return row_.CustomHeight;
                case RowValueKind.Height: return row_.Height;
                case RowValueKind.Hidden: return row_.Hidden;
                case RowValueKind.Merged: return row_.Merged;
                case RowValueKind.OutlineLevel: return row_.OutlineLevel;
                case RowValueKind.PageBreak: return row_.PageBreak;
                case RowValueKind.Phonetic: return row_.Phonetic;
                case RowValueKind.StyleName: return row_.StyleName;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        #endregion
    }
}
