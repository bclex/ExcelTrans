using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    public class CommandCol : IExcelCommand
    {
        public Func<ExcelContext, Collection<string>, object, int, CommandRtn> Func { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        public CommandCol(Func<object, int, CommandRtn> func, params IExcelCommand[] cmds)
            : this((a, b, c, d) => func(c, d), cmds) { }
        public CommandCol(Func<ExcelContext, Collection<string>, object, int, CommandRtn> func, params IExcelCommand[] cmds)
        {
            Func = func;
            Cmds = cmds;
        }
        public CommandCol(Func<object, int, CommandRtn> func, Func<IExcelCommand[]> cmds)
            : this((a, b, c, d) => func(c, d), a => cmds()) { }
        public CommandCol(Func<ExcelContext, Collection<string>, object, int, CommandRtn> func, Func<ExcelContext, IExcelCommand[]> cmds)
        {
            Func = func;
            Cmds = null;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Func = ExcelContext.DecodeFunc<ExcelContext, Collection<string>, object, int, CommandRtn>(r);
            Cmds = ExcelContext.DecodeCommands(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelContext.EncodeFunc(w, Func);
            ExcelContext.EncodeCommands(w, Cmds);
        }

        void IExcelCommand.Execute(ExcelContext ctx) { }
    }
}