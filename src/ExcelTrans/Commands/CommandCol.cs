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
        public CommandCol(Func<object, int, CommandRtn> func, Func<IExcelCommand[]> cmds)
            : this((a, b, c, d) => func(c, d), a => cmds()) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, int, CommandRtn> func, Func<IExcelContext, IExcelCommand[]> cmds)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = null;
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

        void IExcelCommand.Execute(IExcelContext ctx) { }

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CommandCol{(When == When.Normal ? null : $"[{When}]")}: [func]"); ExcelSerDes.DescribeCommands(w, pad, Cmds); }
    }
}