using ExcelTrans.Utils;
using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    public class CommandCol : IExcelCommand
    {
        public When When { get; private set; }
        public Func<IExcelContext, Collection<string>, object, int, CommandRtn> Func { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        public CommandCol(Func<object, int, CommandRtn> func, params IExcelCommand[] cmds)
            : this((a, b, c, d) => func(c, d), cmds) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, int, CommandRtn> func, params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        public CommandCol(Func<object, int, CommandRtn> func, Action command)
            : this((a, b, c, d) => func(c, d), command) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, int, CommandRtn> func, Action command)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        public CommandCol(Func<object, int, CommandRtn> func, Action<IExcelContext> command)
            : this((a, b, c, d) => func(c, d), command) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, int, CommandRtn> func, Action<IExcelContext> command)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Func = ExcelSerDes.DecodeFunc<IExcelContext, Collection<string>, object, int, CommandRtn>(r);
            Cmds = ExcelSerDes.DecodeCommands(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelSerDes.EncodeFunc(w, Func);
            ExcelSerDes.EncodeCommands(w, Cmds);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.CmdCols.Push(this);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CommandCol{(When == When.Normal ? null : $"[{When}]")}: [func]"); ExcelSerDes.DescribeCommands(w, pad, Cmds); }

        internal static void Reset(IExcelContext ctx, int idx)
        {
            while (ctx.CmdRows.Count > idx)
                ctx.CmdRows.Pop();
        }
    }
}