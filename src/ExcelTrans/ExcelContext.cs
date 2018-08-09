using ExcelTrans.Commands;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;

namespace ExcelTrans
{
    public class ExcelContext : IDisposable
    {
        public ExcelContext()
        {
            p = new ExcelPackage();
            wb = p.Workbook;
        }
        public void Dispose() => p.Dispose();

        public int x = 1, xstart = 1, dx = 1;
        public int y = 1, dy = 1;
        public ExcelPackage p;
        public ExcelWorkbook wb;
        public ExcelWorksheet ws;
        public readonly Stack<Tuple<CommandRow[], CommandCol[]>> cmds = new Stack<Tuple<CommandRow[], CommandCol[]>>();
        public readonly Stack<IExcelCommandSet> sets = new Stack<IExcelCommandSet>();
        public static readonly List<object> funcs = new List<object>();

        public ExcelWorksheet EnsureWorksheet() => ws ?? (ws = wb.Worksheets.Add($"Sheet {wb.Worksheets.Count + 1}"));

        public object GetCtx() => new Tuple<int, int>(cmds.Count, sets.Count);
        public void SetCtx(object si)
        {
            var v = (Tuple<int, int>)si;
            PopCommand.Reset(this, v.Item1);
            PopSet.Reset(this, v.Item2);
        }


        // MARK: ENCODE/DECODE

        public static string Encode(params IExcelCommand[] cmds)
        {
            using (var b = new MemoryStream())
            using (var w = new BinaryWriter(b))
            {
                EncodeCommands(w, cmds);
                b.Position = 0;
                return Convert.ToBase64String(b.ToArray());
            }
        }
        public static IExcelCommand[] Decode(string packed)
        {
            using (var b = new MemoryStream(Convert.FromBase64String(packed)))
            using (var r = new BinaryReader(b))
            {
                return DecodeCommands(r);
            }
        }

        public static void EncodeAction(BinaryWriter w, object action) { w.Write((ushort)funcs.Count); funcs.Add(action); }
        public static Action DecodeAction(BinaryReader r) => (Action)funcs[r.ReadUInt16()];
        public static Action<T1> DecodeAction<T1>(BinaryReader r) => (Action<T1>)funcs[r.ReadUInt16()];
        public static Action<T1, T2> DecodeAction<T1, T2>(BinaryReader r) => (Action<T1, T2>)funcs[r.ReadUInt16()];
        public static Action<T1, T2, T3> DecodeAction<T1, T2, T3>(BinaryReader r) => (Action<T1, T2, T3>)funcs[r.ReadUInt16()];
        public static Action<T1, T2, T3, T4> DecodeAction<T1, T2, T3, T4>(BinaryReader r) => (Action<T1, T2, T3, T4>)funcs[r.ReadUInt16()];
        public static Action<T1, T2, T3, T4, T5> DecodeAction<T1, T2, T3, T4, T5>(BinaryReader r) => (Action<T1, T2, T3, T4, T5>)funcs[r.ReadUInt16()];

        public static void EncodeFunc(BinaryWriter w, object func) { w.Write((ushort)funcs.Count); funcs.Add(func); }
        public static Func<TR> DecodeFunc<TR>(BinaryReader r) => (Func<TR>)funcs[r.ReadUInt16()];
        public static Func<T1, TR> DecodeFunc<T1, TR>(BinaryReader r) => (Func<T1, TR>)funcs[r.ReadUInt16()];
        public static Func<T1, T2, TR> DecodeFunc<T1, T2, TR>(BinaryReader r) => (Func<T1, T2, TR>)funcs[r.ReadUInt16()];
        public static Func<T1, T2, T3, TR> DecodeFunc<T1, T2, T3, TR>(BinaryReader r) => (Func<T1, T2, T3, TR>)funcs[r.ReadUInt16()];
        public static Func<T1, T2, T3, T4, TR> DecodeFunc<T1, T2, T3, T4, TR>(BinaryReader r) => (Func<T1, T2, T3, T4, TR>)funcs[r.ReadUInt16()];
        public static Func<T1, T2, T3, T4, T5, TR> DecodeFunc<T1, T2, T3, T4, T5, TR>(BinaryReader r) => (Func<T1, T2, T3, T4, T5, TR>)funcs[r.ReadUInt16()];

        public static void EncodeCommands(BinaryWriter w, IExcelCommand[] cmds)
        {
            if (cmds == null)
            {
                w.Write((ushort)0);
                return;
            }
            w.Write((ushort)cmds.Length);
            foreach (var cmd in cmds)
            {
                var cmdId = ExcelService.cmds.IndexOf(cmd.GetType());
                if (cmdId == -1) throw new InvalidOperationException($"{cmd} does not exist");
                w.Write((ushort)cmdId);
                cmd.Write(w);
            }
        }
        public static IExcelCommand[] DecodeCommands(BinaryReader r)
        {
            var cmds = new IExcelCommand[r.ReadUInt16()];
            for (var i = 0; i < cmds.Length; i++)
            {
                var cmd = (IExcelCommand)FormatterServices.GetUninitializedObject(ExcelService.cmds[r.ReadUInt16()]);
                cmd.Read(r);
                cmds[i] = cmd;
            }
            return cmds;
        }


        // MARK: EXECUTE

        public object Execute(IExcelCommand[] cmds)
        {
            var si = GetCtx();
            foreach (var cmd in cmds)
                cmd.Execute(this);
            return si;
        }

        public CommandRtn ExecuteRow(Collection<string> s)
        {
            var cr = CommandRtn.None;
            foreach (var cmd in cmds.SelectMany(z => z.Item1))
            {
                var r = cmd.Predicate(this, s);
                if ((r & CommandRtn.Execute) == CommandRtn.Execute)
                    SetCtx(Execute(cmd.Cmds));
                cr |= r;
            }
            return cr;
        }

        public CommandRtn ExecuteCol(Collection<string> s, object v, int i)
        {
            var cr = CommandRtn.None;
            foreach (var cmd in cmds.SelectMany(z => z.Item2))
            {
                var r = cmd.Func(this, s, v, i);
                if ((r & CommandRtn.Execute) == CommandRtn.Execute)
                    SetCtx(Execute(cmd.Cmds));
                cr |= r;
            }
            return cr;
        }
    }
}