using ExcelTrans.Commands;
using ExcelTrans.Services;
using OfficeOpenXml;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExcelTrans
{
    public interface IExcelCommand
    {
        void Read(BinaryReader r);
        void Write(BinaryWriter w);
        void Execute(IExcelContext ctx);
    }

    public interface IExcelCommandSet
    {
        void Add(Collection<string> s);
        void Execute(IExcelContext ctx);
    }

    public static class ExcelService
    {
        public static readonly string Stream = "^q=";
        public static readonly string Break = "^q!";
        public static string PopSet => $"{Stream}{ExcelSerDes.Encode(new PopSet())}";
        public static string Encode(params IExcelCommand[] cmds) => $"{Stream}{ExcelSerDes.Encode(cmds)}";
        public static IExcelCommand[] Decode(string packed) => ExcelSerDes.Decode(packed.Substring(Stream.Length));

        public static Tuple<Stream, string, string> Transform(Tuple<Stream, string, string> a)
        {
            using (var s1 = a.Item1)
            {
                s1.Seek(0, SeekOrigin.Begin);
                var sr = new StreamReader(s1);
                var s2 = new MemoryStream(Build(sr));
                return new Tuple<Stream, string, string>(s2, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", a.Item3.Replace(".csv", ".xlsx"));
            }
        }

        static byte[] Build(TextReader sr)
        {
            var cr = new CsvReader();
            using (var ctx = new ExcelContext())
            {
                cr.Execute(sr, x =>
                {
                    ctx.CsvY++;
                    if (x == null || x.Count == 0) return true;
                    else if (x[0].StartsWith(Stream)) return ctx.Execute(Decode(x[0])) != null;
                    else if (x[0].StartsWith(Break)) return false;
                    if (ctx.Sets.Count == 0) ctx.WriteRow(x);
                    else ctx.Sets.Peek().Add(x);
                    return true;
                }).Any(x => !x);
                return ctx.P.GetAsByteArray();
            }
        }

        public static string GetAddress(int row, int column) => ExcelCellBase.GetAddress(row, column);
        public static string GetAddress(int row, bool absoluteRow, int column, bool absoluteCol) => ExcelCellBase.GetAddress(row, absoluteRow, column, absoluteCol);
        public static string GetAddress(int row, int column, bool absolute) => ExcelCellBase.GetAddress(row, column, absolute);
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn);
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn, bool absolute) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, absolute);
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn, bool fixedFromRow, bool fixedFromColumn, bool fixedToRow, bool fixedToColumn) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, fixedFromRow, fixedFromColumn, fixedToRow, fixedToColumn);
        public static string GetAddress(Address r, int row, int col) => $"^{(int)r}:{row}:{col}";
        public static string GetAddress(Address r, int fromRow, int fromColumn, int toRow, int toColumn) => $"^{(int)r}:{fromRow}:{fromColumn}:{toRow}:{toColumn}";
        public static string GetAddress(IExcelContext ctx, Address r, int row, int col) => DecodeAddress(ctx, GetAddress(r, row, col));
        public static string GetAddress(IExcelContext ctx, Address r, int fromRow, int fromColumn, int toRow, int toColumn) => DecodeAddress(ctx, GetAddress(r, fromRow, fromColumn, toRow, toColumn));
        public static string DecodeAddress(IExcelContext ctx, string address)
        {
            if (!address.StartsWith("^")) return address;
            string r;
            var vec = address.Substring(1).Split(':').Select(x => int.Parse(x)).ToArray();
            if (vec.Length == 3)
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.Cell: r = ExcelCellBase.GetAddress(ctx.Y + vec[1], ctx.X + vec[2]); break;
                    case Address.Range: r = ExcelCellBase.GetAddress(ctx.Y, ctx.X, ctx.Y + vec[1], ctx.X + vec[2]); break;
                    default: throw new ArgumentOutOfRangeException("address");
                }
            else if (vec.Length == 5)
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.Range: r = ExcelCellBase.GetAddress(ctx.Y + vec[1], ctx.X + vec[2], ctx.Y + vec[3], ctx.X + vec[4]); break;
                    default: throw new ArgumentOutOfRangeException("address");
                }
            else throw new ArgumentOutOfRangeException("address");
            return r;
        }

        public static object ParseValue(string v) =>
            DateTime.TryParse(v, out var vd) ? vd :
                long.TryParse(v, out var vl) ? vl :
                float.TryParse(v, out var vf) ? (object)vf :
                v;
    }
}
