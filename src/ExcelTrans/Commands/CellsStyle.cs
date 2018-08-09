using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct CellsStyle : IExcelCommand
    {
        public string Cells { get; private set; }
        public string[] Styles { get; private set; }

        public CellsStyle(int row, int col, params string[] styles)
            : this(ExcelCellBase.GetAddress(row, col), styles) { }
        public CellsStyle(int fromRow, int fromCol, int toRow, int toCol, params string[] styles)
            : this(ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol), styles) { }
        public CellsStyle(ExcelContext r, int plusRow, int plusCol, params string[] styles)
            : this(ExcelCellBase.GetAddress(r.y, r.x, r.y + plusRow, r.x + plusCol), styles) { }
        public CellsStyle(string cells, params string[] styles)
        {
            Cells = cells;
            Styles = styles;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            Styles = new string[r.ReadUInt16()];
            for (var i = 0; i < Styles.Length; i++)
                Styles[i] = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write((ushort)Styles.Length);
            for (var i = 0; i < Styles.Length; i++)
                w.Write(Styles[i]);
        }

        // https://stackoverflow.com/questions/40209636/epplus-number-format/40214134
        void IExcelCommand.Execute(ExcelContext ctx)
        {
            var range = ctx.ws.Cells[Cells];
            foreach (var style in Styles)
            {
                // number-format
                if (style.StartsWith("n"))
                {
                    if (style.StartsWith("n:")) range.Style.Numberformat.Format = style.Substring(2);
                    else if (style == "n$") range.Style.Numberformat.Format = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \" - \"??_);_(@_)"; // "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
                    else if (style == "n%") range.Style.Numberformat.Format = "0%";
                    else if (style == "n,") range.Style.Numberformat.Format = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)";
                    else throw new InvalidOperationException($"{style}");
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
                    else throw new InvalidOperationException($"{style}");
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
                    else throw new InvalidOperationException($"{style}");
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
                else throw new InvalidOperationException($"{style}");
            }
        }
    }
}