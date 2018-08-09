using ExcelTrans.Commands;
using ExcelTrans.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExcelTrans
{
    public interface IExcelCommand
    {
        void Read(BinaryReader r);
        void Write(BinaryWriter w);
        void Execute(ExcelContext ctx);
    }

    public interface IExcelCommandSet
    {
        void Add(Collection<string> s);
        void Execute(ExcelContext ctx);
    }

    public static class ExcelService
    {
        public static readonly List<Type> cmds = new List<Type>() {
            typeof(CellsStyle), typeof(CellsValue), typeof(Command), typeof(CommandCol), typeof(CommandRow),
            typeof(PopCommand), typeof(PopSet), typeof(PushCommand), typeof(PushSet), typeof(WorksheetsAdd) };
        public static readonly string Stream = "^q=";
        public static readonly string Break = "^q!";
        public static string PopSet => $"{Stream}{ExcelContext.Encode(new PopSet())}";
        public static string Encode(params IExcelCommand[] cmds) => $"{Stream}{ExcelContext.Encode(cmds)}";
        public static IExcelCommand[] Decode(string packed) => ExcelContext.Decode(packed.Substring(Stream.Length));

        public static Tuple<Stream, string, string> Transform(Tuple<Stream, string, string> a)
        {
            using (var s = a.Item1)
            {
                s.Seek(0, SeekOrigin.Begin);
                var sr = new StreamReader(s);
                var newS = new MemoryStream(Build(sr));
                return new Tuple<Stream, string, string>(newS, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", a.Item3.Replace(".csv", ".xlsx"));
            }
        }

        static byte[] Build(TextReader sr)
        {
            var cr = new CsvReader();
            using (var ctx = new ExcelContext())
            {
                cr.Execute(sr, x =>
                {
                    if (x == null || x.Count == 0) return true;
                    else if (x[0].StartsWith(Stream)) return ctx.Execute(Decode(x[0])) != null;
                    else if (x[0].StartsWith(Break)) return false;
                    if (ctx.sets.Count == 0) ProcessRow(ctx, x);
                    else ctx.sets.Peek().Add(x);
                    return true;
                }).Any(x => !x);
                return ctx.p.GetAsByteArray();
            }
        }

        public static bool ProcessRow(ExcelContext ctx, Collection<string> s)
        {
            ctx.EnsureWorksheet();
            var ws = ctx.ws;
            // run
            var cr = ctx.ExecuteRow(s);
            if ((cr & CommandRtn.Skip) == CommandRtn.Skip)
                return true;
            ctx.x = ctx.xstart;
            for (var i = 0; i < s.Count; i++)
            {
                var v = s[i];
                var y = long.TryParse(v, out var vl) ? vl :
                    float.TryParse(v, out var vf) ? (object)vf :
                    v;
                // run
                cr = ctx.ExecuteCol(s, v, i);
                if ((cr & CommandRtn.Skip) == CommandRtn.Skip)
                    continue;
                if ((cr & CommandRtn.Formula) != CommandRtn.Formula) ws.SetValue(ctx.y, ctx.x, y);
                else ws.Cells[ctx.y, ctx.x].Formula = v;
                ctx.x += ctx.dx;
            }
            ctx.y += ctx.dy;
            return true;
        }
    }
}
