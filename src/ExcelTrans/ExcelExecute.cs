using ExcelTrans.Commands;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace ExcelTrans
{
    public static class ExcelExecute
    {
        public static object Execute(this IExcelContext ctx, IExcelCommand[] cmds)
        {
            var si = ctx.GetCtx();
            foreach (var cmd in cmds)
                cmd.Execute(ctx);
            return si;
        }

        public static CommandRtn ExecuteRow(this IExcelContext ctx, WhenRow when, Collection<string> s)
        {
            var cr = CommandRtn.None;
            foreach (var cmd in ctx.Cmds.SelectMany(z => z.Item1.Where(x => (x.When & when) == when)))
            {
                var r = cmd.Func(ctx, s);
                if ((r & CommandRtn.Execute) == CommandRtn.Execute)
                    ctx.SetCtx(ctx.Execute(cmd.Cmds));
                cr |= r;
            }
            return cr;
        }

        public static CommandRtn ExecuteCol(this IExcelContext ctx, Collection<string> s, object v, int i)
        {
            var cr = CommandRtn.None;
            foreach (var cmd in ctx.Cmds.SelectMany(z => z.Item2))
            {
                var r = cmd.Func(ctx, s, v, i);
                if ((r & CommandRtn.Execute) == CommandRtn.Execute)
                    ctx.SetCtx(ctx.Execute(cmd.Cmds));
                cr |= r;
            }
            return cr;
        }

        public static void WriteFirst(this IExcelContext ctx, Collection<string> s)
        {
            ctx.ExecuteRow(WhenRow.FirstRow, s);
        }

        public static void WriteRow(this IExcelContext ctx, Collection<string> s)
        {
            var ws = ((ExcelContext)ctx).EnsureWorksheet();
            // execute-row-before
            var cr = ctx.ExecuteRow(WhenRow.BeforeRow, s);
            if ((cr & CommandRtn.Skip) == CommandRtn.Skip)
                return;
            ctx.X = ctx.XStart;
            for (var i = 0; i < s.Count; i++)
            {
                ctx.CsvX = i + 1;
                var v = ExcelService.ParseValue(s[i]);
                // execute-col
                cr = ctx.ExecuteCol(s, v, i);
                if ((cr & CommandRtn.Skip) == CommandRtn.Skip)
                    continue;
                if ((cr & CommandRtn.Formula) != CommandRtn.Formula) ws.SetValue(ctx.Y, ctx.X, v);
                else ws.Cells[ctx.Y, ctx.X].Formula = s[i];
                if (v is DateTime) ws.Cells[ExcelCellBase.GetAddress(ctx.Y, ctx.X)].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                ctx.X += ctx.DeltaX;
            }
            ctx.Y += ctx.DeltaY;
            // execute-row-after
            cr = ctx.ExecuteRow(WhenRow.AfterRow, s);
        }

        public static void WriteLast(this IExcelContext ctx, Collection<string> s)
        {
            ctx.ExecuteRow(WhenRow.LastRow, s);
        }

        // https://stackoverflow.com/questions/40209636/epplus-number-format/40214134
        public static void CellsStyle(this IExcelContext ctx, int row, int col, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(row, col), styles);
        public static void CellsStyle(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(r, 0, 0), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, int row, int col, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(r, row, col), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] styles) => ctx.CellsStyle(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), styles);
        public static void CellsStyle(this IExcelContext ctx, string cells, string[] styles)
        {
            var range = ((ExcelContext)ctx).WS.Cells[ExcelService.DecodeAddress(ctx, cells)];
            foreach (var style in styles)
            {
                // number-format
                if (style.StartsWith("n"))
                {
                    if (style.StartsWith("n:")) range.Style.Numberformat.Format = style.Substring(2);
                    else if (style == "n$") range.Style.Numberformat.Format = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \" - \"??_);_(@_)"; // "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
                    else if (style == "n%") range.Style.Numberformat.Format = "0%";
                    else if (style == "n,") range.Style.Numberformat.Format = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)";
                    else if (style == "nd") range.Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    else throw new InvalidOperationException($"{style} not defined");
                }
                // font
                else if (style.StartsWith("f"))
                {
                    if (style.StartsWith("f:")) range.Style.Font.Name = style.Substring(2);
                    else if (style.StartsWith("fx")) range.Style.Font.Size = float.Parse(style.Substring(2));
                    else if (style.StartsWith("ff")) range.Style.Font.Family = int.Parse(style.Substring(2));
                    //else if (style.StartsWith("fc")) range.Style.Font.Color = int.Parse(style.Substring(2));
                    else if (style.StartsWith("fs")) range.Style.Font.Scheme = style.Substring(2);
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
                //else if (style.StartsWith("l"))
                //{
                //    range.Style.Fill
                //}
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

        public static void CellsValue(this IExcelContext ctx, int row, int col, object value, bool formula = false) => ctx.CellsValue(ExcelService.GetAddress(row, col), value, formula);
        public static void CellsValue(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, object value, bool formula = false) => ctx.CellsValue(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, formula);
        public static void CellsValue(this IExcelContext ctx, Address r, object value, bool formula = false) => ctx.CellsValue(ExcelService.GetAddress(r, 0, 0), value, formula);
        public static void CellsValue(this IExcelContext ctx, Address r, int row, int col, object value, bool formula = false) => ctx.CellsValue(ExcelService.GetAddress(r, row, col), value, formula);
        public static void CellsValue(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, object value, bool formula = false) => ctx.CellsValue(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, formula);
        public static void CellsValue(this IExcelContext ctx, string cells, object value, bool formula = false)
        {
            var range = ((ExcelContext)ctx).WS.Cells[ExcelService.DecodeAddress(ctx, cells)];
            if (!formula) range.Value = value;
            else range.Formula = (string)value;
            if (value is DateTime) range.Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
        }

        public static object GetCellsValue(this IExcelContext ctx, int row, int col, bool formula = false) => ctx.GetCellsValue(ExcelService.GetAddress(row, col), formula);
        public static object GetCellsValue(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, bool formula = false) => ctx.GetCellsValue(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), formula);
        public static object GetCellsValue(this IExcelContext ctx, Address r, bool formula = false) => ctx.GetCellsValue(ExcelService.GetAddress(r, 0, 0), formula);
        public static object GetCellsValue(this IExcelContext ctx, Address r, int row, int col, bool formula = false) => ctx.GetCellsValue(ExcelService.GetAddress(r, row, col), formula);
        public static object GetCellsValue(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, bool formula = false) => ctx.GetCellsValue(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), formula);
        public static object GetCellsValue(this IExcelContext ctx, string cells, bool formula = false)
        {
            var range = ((ExcelContext)ctx).WS.Cells[ExcelService.DecodeAddress(ctx, cells)];
            if (!formula) return range.Value;
            else return range.Formula;
        }
    }
}
