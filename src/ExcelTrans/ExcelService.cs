using ExcelTrans.Commands;
using ExcelTrans.Services;
using ExcelTrans.Utils;
using OfficeOpenXml;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExcelTrans
{
    public interface IExcelCommand
    {
        When When { get; }
        void Read(BinaryReader r);
        void Write(BinaryWriter w);
        void Execute(IExcelContext ctx);
        void Describe(StringWriter w, int pad);
    }

    public interface IExcelCommandSet
    {
        void Add(Collection<string> s);
        void Execute(IExcelContext ctx);
    }

    public static class ExcelService
    {
        public static readonly string Comment = "^q|";
        public static readonly string Stream = "^q=";
        public static readonly string Break = "^q!";
        public static string PopCommand => $"{Stream}{ExcelSerDes.Encode(new PopCommand())}";
        public static string PopSet => $"{Stream}{ExcelSerDes.Encode(new PopSet())}";
        public static string Encode(bool describe, params IExcelCommand[] cmds) => $"{(describe ? ExcelSerDes.Describe(Comment, cmds) : null)}{Stream}{ExcelSerDes.Encode(cmds)}";
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
                    if (x == null || x.Count == 0 || x[0].StartsWith(Comment)) return true;
                    else if (x[0].StartsWith(Stream)) { var r = ctx.Execute(Decode(x[0]), out var after) != null; after?.Invoke(); return r; }
                    else if (x[0].StartsWith(Break)) return false;
                    ctx.CsvY++;
                    if (ctx.Sets.Count == 0) ctx.WriteRow(x);
                    else ctx.Sets.Peek().Add(x);
                    return true;
                }).Any(x => !x);
                return ctx.P.GetAsByteArray();
            }
        }

        public static string GetAddressCol(int column) => ExcelCellBase.GetAddressCol(column);
        public static string GetAddressRow(int row) => ExcelCellBase.GetAddressRow(row);
        public static string GetAddress(int row, string column) => ExcelCellBase.GetAddress(row, ColToInt(column));
        public static string GetAddress(int row, int column) => ExcelCellBase.GetAddress(row, column);
        public static string GetAddress(int row, bool absoluteRow, string column, bool absoluteCol) => ExcelCellBase.GetAddress(row, absoluteRow, ColToInt(column), absoluteCol);
        public static string GetAddress(int row, bool absoluteRow, int column, bool absoluteCol) => ExcelCellBase.GetAddress(row, absoluteRow, column, absoluteCol);
        public static string GetAddress(int row, string column, bool absolute) => ExcelCellBase.GetAddress(row, ColToInt(column), absolute);
        public static string GetAddress(int row, int column, bool absolute) => ExcelCellBase.GetAddress(row, column, absolute);
        public static string GetAddress(int fromRow, string fromColumn, int toRow, string toColumn) => ExcelCellBase.GetAddress(fromRow, ColToInt(fromColumn), toRow, ColToInt(toColumn));
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn);
        public static string GetAddress(int fromRow, string fromColumn, int toRow, string toColumn, bool absolute) => ExcelCellBase.GetAddress(fromRow, ColToInt(fromColumn), toRow, ColToInt(toColumn), absolute);
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn, bool absolute) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, absolute);
        public static string GetAddress(int fromRow, string fromColumn, int toRow, string toColumn, bool fixedFromRow, bool fixedFromColumn, bool fixedToRow, bool fixedToColumn) => ExcelCellBase.GetAddress(fromRow, ColToInt(fromColumn), toRow, ColToInt(toColumn), fixedFromRow, fixedFromColumn, fixedToRow, fixedToColumn);
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn, bool fixedFromRow, bool fixedFromColumn, bool fixedToRow, bool fixedToColumn) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, fixedFromRow, fixedFromColumn, fixedToRow, fixedToColumn);
        public static string GetAddress(Address r, int row, string col) => $"^{(int)r}:{row}:{ColToInt(col)}";
        public static string GetAddress(Address r, int row, int col) => $"^{(int)r}:{row}:{col}";
        public static string GetAddress(Address r, int fromRow, string fromColumn, int toRow, string toColumn) => $"^{(int)r}:{fromRow}:{ColToInt(fromColumn)}:{toRow}:{ColToInt(toColumn)}";
        public static string GetAddress(Address r, int fromRow, int fromColumn, int toRow, int toColumn) => $"^{(int)r}:{fromRow}:{fromColumn}:{toRow}:{toColumn}";
        public static string GetAddress(IExcelContext ctx, Address r, int row, string col) => DecodeAddress(ctx, GetAddress(r, row, col));
        public static string GetAddress(IExcelContext ctx, Address r, int row, int col) => DecodeAddress(ctx, GetAddress(r, row, col));
        public static string GetAddress(IExcelContext ctx, Address r, int fromRow, string fromColumn, int toRow, string toColumn) => DecodeAddress(ctx, GetAddress(r, fromRow, fromColumn, toRow, toColumn));
        public static string GetAddress(IExcelContext ctx, Address r, int fromRow, int fromColumn, int toRow, int toColumn) => DecodeAddress(ctx, GetAddress(r, fromRow, fromColumn, toRow, toColumn));
        public static string DecodeAddress(IExcelContext ctx, string address)
        {
            if (!address.StartsWith("^")) return address;
            var vec = address.Substring(1).Split(':').Select(x => int.Parse(x)).ToArray();
            if (vec.Length == 3)
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.Cell: return ExcelCellBase.GetAddress(ctx.Y + vec[1], ctx.X + vec[2]);
                    case Address.Range: return ExcelCellBase.GetAddress(ctx.Y, ctx.X, ctx.Y + vec[1], ctx.X + vec[2]);
                    default: throw new ArgumentOutOfRangeException(nameof(address));
                }
            else if (vec.Length == 5)
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.Range: return ExcelCellBase.GetAddress(ctx.Y + vec[1], ctx.X + vec[2], ctx.Y + vec[3], ctx.X + vec[4]);
                    default: throw new ArgumentOutOfRangeException(nameof(address));
                }
            else throw new ArgumentOutOfRangeException(nameof(address));
        }
        internal static string DescribeAddress(string address)
        {
            if (!address.StartsWith("^")) return address;
            var vec = address.Substring(1).Split(':').Select(x => int.Parse(x)).ToArray();
            if (vec.Length == 3)
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.Cell: return $"r:{vec[1]}.{vec[2]}";
                    case Address.Range: return $"r:0.0:{vec[1]}.{vec[2]}";
                    default: throw new ArgumentOutOfRangeException(nameof(address));
                }
            else if (vec.Length == 5)
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.Range: return $"r:{vec[1]}.{vec[2]}:{vec[3]}.{vec[4]}";
                    default: throw new ArgumentOutOfRangeException(nameof(address));
                }
            else throw new ArgumentOutOfRangeException(nameof(address));
        }

        internal static object ParseValue(this string v) =>
            v == null ? null :
                v.Contains("/") && DateTime.TryParse(v, out var vd) ? vd :
                v.Contains(".") && double.TryParse(v, out var vf) ? vf :
                long.TryParse(v, out var vl) ? vl :
                (object)v;

        internal static object CastValue(this string v, Type type) => type == null ? v : Convert.ChangeType(v, type);
        internal static T CastValue<T>(this object v) => (T)Convert.ChangeType(v, typeof(T));

        internal static int RowToInt(string row) => int.Parse(row);
        internal static int ColToInt(string col) => col.ToUpperInvariant().Aggregate(0, (a, x) => (a * 26) + (x - '@'));
    }
}
