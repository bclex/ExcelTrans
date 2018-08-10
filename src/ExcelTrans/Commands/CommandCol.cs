using ExcelTrans.Services;
using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    public class CommandCol : IExcelCommand
    {
        public Func<IExcelContext, Collection<string>, object, int, CommandRtn> Func { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        public CommandCol(Func<object, int, CommandRtn> func, params IExcelCommand[] cmds)
            : this((a, b, c, d) => func(c, d), cmds) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, int, CommandRtn> func, params IExcelCommand[] cmds)
        {
            Func = func;
            Cmds = cmds;
        }
        public CommandCol(Func<object, int, CommandRtn> func, Func<IExcelCommand[]> cmds)
            : this((a, b, c, d) => func(c, d), a => cmds()) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, int, CommandRtn> func, Func<IExcelContext, IExcelCommand[]> cmds)
        {
            Func = func;
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
    }
}